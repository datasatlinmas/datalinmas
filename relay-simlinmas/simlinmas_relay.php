<?php
/**
 * SIMLINMAS Relay - Powerfull Cloudflare-Resilient Version & Test Login
 * Fungsi: relay komunikasi ke SIMLINMAS (pass-through) dengan header & cURL hardening.
 */
declare(strict_types=1);
header('Content-Type: application/json; charset=utf-8');
$origin = $_SERVER['HTTP_ORIGIN'] ?? '*';
header('Access-Control-Allow-Origin: ' . $origin);
header('Vary: Origin');
header('Access-Control-Allow-Methods: POST, OPTIONS');
header('Access-Control-Allow-Headers: Content-Type, X-Relay-Token');
if ($_SERVER['REQUEST_METHOD'] === 'OPTIONS') { http_response_code(204); exit; }

const RELAY_TOKEN        = 'jaUIJ21PG8qJRKeegiyZEFywBApjcMeW3LWVMJ61iI0';
const SIMLINMAS_BASE     = 'https://simlinmas.kemendagri.go.id';
const SIMLINMAS_EMAIL    = 'admin.kabupaten-ponorogo@gmail.com';
const SIMLINMAS_PASSWORD = '123456';
const TIMEOUT            = 120;

function out(array $data, int $code = 200): void {
    http_response_code($code);
    echo json_encode($data, JSON_UNESCAPED_UNICODE | JSON_UNESCAPED_SLASHES);
    exit;
}

function get_token(): string {
    $h = '';
    if (function_exists('getallheaders')) {
        foreach (getallheaders() as $k => $v) {
            if (strtolower((string)$k) === 'x-relay-token') { $h = trim((string)$v); break; }
        }
    }
    if ($h === '') $h = trim((string)($_SERVER['HTTP_X_RELAY_TOKEN'] ?? ''));
    return $h !== '' ? $h : trim((string)($_POST['token'] ?? ''));
}

if ($_SERVER['REQUEST_METHOD'] !== 'POST') out(['ok' => false, 'message' => 'Method harus POST'], 405);
if (!hash_equals(RELAY_TOKEN, get_token())) out(['ok' => false, 'message' => 'Token tidak valid'], 403);

function curl_base(): array {
    $cookie = tempnam(sys_get_temp_dir(), 'simrelay_');
    $ch = curl_init();
    
    // Konfigurasi cURL yang di-hardened agar menyerupai browser modern
    curl_setopt_array($ch, [
        CURLOPT_RETURNTRANSFER => true,
        CURLOPT_FOLLOWLOCATION => true,
        CURLOPT_MAXREDIRS      => 5,
        CURLOPT_TIMEOUT        => TIMEOUT,
        CURLOPT_CONNECTTIMEOUT => 30,
        CURLOPT_COOKIEJAR      => $cookie,
        CURLOPT_COOKIEFILE     => $cookie,
        CURLOPT_USERAGENT      => 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36',
        CURLOPT_SSL_VERIFYPEER => true,
        CURLOPT_SSL_VERIFYHOST => 2,
        CURLOPT_ENCODING       => 'gzip, deflate',
    ]);
    return [$ch, $cookie];
}

function get_default_headers(string $referer = ''): array {
    $headers = [
        'Accept: text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8',
        'Accept-Language: id-ID,id;q=0.9,en-US;q=0.8,en;q=0.7',
        'Cache-Control: max-age=0',
        'Sec-Ch-Ua: "Chromium";v="122", "Not(A:Brand";v="8", "Google Chrome";v="122"',
        'Sec-Ch-Ua-Mobile: ?0',
        'Sec-Ch-Ua-Platform: "Windows"',
        'Sec-Fetch-Dest: document',
        'Sec-Fetch-Mode: navigate',
        'Sec-Fetch-Site: same-origin',
        'Sec-Fetch-User: ?1',
        'Upgrade-Insecure-Requests: 1',
    ];
    if ($referer !== '') {
        $headers[] = 'Referer: ' . $referer;
    }
    return $headers;
}

function curl_req($ch, string $method, string $url, $fields = null, array $customHeaders = []): array {
    curl_setopt($ch, CURLOPT_URL, $url);
    curl_setopt($ch, CURLOPT_CUSTOMREQUEST, $method);
    curl_setopt($ch, CURLOPT_HTTPHEADER, $customHeaders);
    
    if ($method === 'POST') {
        curl_setopt($ch, CURLOPT_POST, true);
        curl_setopt($ch, CURLOPT_POSTFIELDS, $fields);
    } else {
        curl_setopt($ch, CURLOPT_HTTPGET, true);
    }
    
    $body   = curl_exec($ch);
    $status = (int)curl_getinfo($ch, CURLINFO_HTTP_CODE);
    $err    = curl_error($ch);
    
    return [$status, $body === false ? '' : (string)$body, $err];
}

function is_cloudflare_blocked(string $body, int $status): bool {
    if ($status === 403 || $status === 503) {
        if (stripos($body, 'cf-browser-verification') !== false || 
            stripos($body, 'cloudflare') !== false || 
            stripos($body, 'Attention Required!') !== false ||
            stripos($body, 'Just a moment...') !== false) {
            return true;
        }
    }
    return false;
}

function get_csrf(string $html): string {
    if (preg_match('/<meta\s+name="csrf-token"\s+content="([^"]+)"/i', $html, $m)) return $m[1];
    if (preg_match('/name="_token"\s+value="([^"]+)"/i', $html, $m)) return $m[1];
    return '';
}

function sim_login($ch): array {
    $loginUrl = SIMLINMAS_BASE . '/login';
    [$s1, $html, $e1] = curl_req($ch, 'GET', $loginUrl, null, get_default_headers(SIMLINMAS_BASE . '/'));
    
    if ($e1) return [false, 'Gagal buka login (cURL Error): ' . $e1];
    if (is_cloudflare_blocked($html, $s1)) return [false, 'Diblokir Cloudflare saat mengakses halaman login (IP Terbanned/Challenge).'];
    
    $csrf = get_csrf($html);
    if ($csrf === '') return [false, 'CSRF login tidak ditemukan di dalam HTML.'];
    
    $postFields = http_build_query([
        '_token'   => $csrf,
        'email'    => SIMLINMAS_EMAIL,
        'password' => SIMLINMAS_PASSWORD,
    ]);
    
    $headers = get_default_headers($loginUrl);
    $headers[] = 'Content-Type: application/x-www-form-urlencoded';
    $headers[] = 'Origin: ' . SIMLINMAS_BASE;
    
    [$s2, $body2, $e2] = curl_req($ch, 'POST', $loginUrl, $postFields, $headers);
    if ($e2) return [false, 'Gagal mengirim data login: ' . $e2];
    if (is_cloudflare_blocked($body2, $s2)) return [false, 'Diblokir Cloudflare saat POST login.'];
    
    [$s3, $html3, $e3] = curl_req($ch, 'GET', SIMLINMAS_BASE . '/linmas', null, get_default_headers($loginUrl));
    if ($e3) return [false, 'Gagal cek sesi dashboard: ' . $e3];
    if ($s3 >= 400 || is_cloudflare_blocked($html3, $s3)) {
        return [false, 'Login gagal atau sesi ditolak Cloudflare (HTTP status: ' . $s3 . ')'];
    }
    
    return [true, 'ok'];
}

function sim_list_nik($ch): array {
    $niks  = [];
    $start = 0;
    $total = 0;
    $debug = '';
    
    $ajaxHeaders = [
        'Accept: application/json, text/javascript, */*; q=0.01',
        'X-Requested-With: XMLHttpRequest',
        'Referer: ' . SIMLINMAS_BASE . '/linmas',
    ];
    
    [$st, $body, $err] = curl_req($ch, 'GET', SIMLINMAS_BASE . '/linmas?start=0&length=1', null, $ajaxHeaders);
    if ($err) { 
        $debug = 'err: ' . $err; 
    } else {
        if (is_cloudflare_blocked($body, $st)) {
            return [false, 'Cloudflare memblokir akses data list NIK', [], 0, 'Cloudflare Blocked'];
        }
        $j = json_decode($body, true);
        if (is_array($j)) {
            $total = (int)($j['recordsFiltered'] ?? $j['recordsTotal'] ?? 0);
        } else {
            $debug = 'status=' . $st . ' body=' . substr($body, 0, 300);
        }
    }
    
    while ($start <= ($total > 0 ? $total + 400 : 5000)) {
        [$st, $body, $err] = curl_req($ch, 'GET', SIMLINMAS_BASE . '/linmas?start=' . $start . '&length=200', null, $ajaxHeaders);
        if ($err) return [false, 'Gagal ambil NIK start=' . $start . ': ' . $err, $niks, $total, $debug];
        
        if (is_cloudflare_blocked($body, $st)) {
            return [false, 'Cloudflare blokir pada paginasi NIK start=' . $start, $niks, $total, 'Cloudflare Blocked'];
        }
        
        $j = json_decode($body, true);
        $rows = is_array($j) ? ($j['data'] ?? []) : [];
        if (!is_array($rows) || count($rows) === 0) { 
            if (!$debug) $debug = 'status=' . $st . ' body=' . substr($body, 0, 300); 
            break; 
        }
        
        foreach ($rows as $r) {
            $nk = trim((string)($r['nik'] ?? ''));
            $id = trim((string)($r['id'] ?? ''));
            $nr = trim((string)($r['nomor_registrasi'] ?? ''));
            if ($nk !== '') $niks[$nk] = ['id'=>$id,'nomor_registrasi'=>$nr];
        }
        
        $start += 200;
        if ($total === 0 && is_array($j)) {
            $total = (int)($j['recordsFiltered'] ?? $j['recordsTotal'] ?? 0);
        }
    }
    $nikList = array_keys($niks);
    return [true, 'ok', $nikList, $total, $debug, $niks];
}

$action = trim((string)($_POST['action'] ?? 'insert'));
[$ch, $cookie] = curl_base();

try {
    // TAMBAHAN AKSI: test_login (untuk uji coba khusus cek koneksi & cloudflare)
    if ($action === 'test_login') {
        [$okLogin, $msgLogin] = sim_login($ch);
        if (!$okLogin) {
            out([
                'ok' => false,
                'action' => 'test_login',
                'message' => 'GAGAL LOGIN SIMLINMAS: ' . $msgLogin,
                'suggestion' => 'Kemungkinan diblokir Cloudflare atau kredensial salah.'
            ], 200);
        }
        
        [$s3, $htmlTambah, $e3] = curl_req($ch, 'GET', SIMLINMAS_BASE . '/linmas/tambah', null, get_default_headers(SIMLINMAS_BASE . '/linmas'));
        $csrfFound = (get_csrf($htmlTambah) !== '');

        out([
            'ok' => true,
            'action' => 'test_login',
            'message' => 'SUKSES! Relay berhasil login ke SIMLINMAS dan menembus proteksi Cloudflare.',
            'details' => [
                'login_status' => 'Berhasil',
                'akses_form_tambah' => ($s3 === 200 && $csrfFound) ? 'Normal & CSRF Didapat' : 'Bermasalah (HTTP: '.$s3.')',
                'csrf_token_length' => strlen(get_csrf($htmlTambah))
            ]
        ], 200);
    }

    [$okLogin, $msgLogin] = sim_login($ch);
    if (!$okLogin) out(['ok' => false, 'action' => $action, 'message' => $msgLogin], 502);

    if ($action === 'ping') {
        out(['ok' => true, 'action' => 'ping', 'message' => 'Login SIMLINMAS OK & Lolos Cloudflare']);
    }

    if ($action === 'list_nik') {
        [$ok, $msg, $niks, $total, $debug] = sim_list_nik($ch);
        if (!$ok) out(['ok' => false, 'action' => 'list_nik', 'message' => $msg, 'niks' => $niks, 'total' => $total, 'debug' => $debug], 502);
        out(['ok' => true, 'action' => 'list_nik', 'total' => $total, 'count' => count($niks), 'niks' => $niks, 'debug' => $debug]);
    }

    function wilayah_map(): array { static $m = null; if ($m===null) $m = json_decode('{"babadan":{"id":"350216","nama":"Babadan","villages":{"babadan":{"id":"3502162013","nama":"Babadan"},"bareng":{"id":"3502162008","nama":"Bareng"},"cekok":{"id":"3502162002","nama":"Cekok"},"gupolo":{"id":"3502162006","nama":"Gupolo"},"japan":{"id":"3502162004","nama":"Japan"},"kadipaten":{"id":"3502161005","nama":"Kadipaten"},"kertosari":{"id":"3502161001","nama":"Kertosari"},"lembah":{"id":"3502162011","nama":"Lembah"},"ngunut":{"id":"3502162009","nama":"Ngunut"},"patihan wetan":{"id":"3502161003","nama":"Patihan Wetan"},"polorejo":{"id":"3502162007","nama":"Polorejo"},"pondok":{"id":"3502162012","nama":"Pondok"},"purwosari":{"id":"3502162014","nama":"Purwosari"},"sukosari":{"id":"3502162010","nama":"Sukosari"},"trisono":{"id":"3502162015","nama":"Trisono"}}},"badegan":{"id":"350213","nama":"Badegan","villages":{"badegan":{"id":"3502132010","nama":"Badegan"},"bandaralim":{"id":"3502132006","nama":"Bandaralim"},"biting":{"id":"3502132009","nama":"Biting"},"dayakan":{"id":"3502132001","nama":"Dayakan"},"kapuran":{"id":"3502132007","nama":"Kapuran"},"karangan":{"id":"3502132002","nama":"Karangan"},"karangjoho":{"id":"3502132004","nama":"Karangjoho"},"tanjunggunung":{"id":"3502132003","nama":"Tanjunggunung"},"tanjungrejo":{"id":"3502132005","nama":"Tanjungrejo"},"watubonang":{"id":"3502132008","nama":"Watubonang"}}},"balong":{"id":"350211","nama":"Balong","villages":{"bajang":{"id":"3502112012","nama":"Bajang"},"balong":{"id":"3502112013","nama":"Balong"},"bulak":{"id":"3502112003","nama":"Bulak"},"bulu kidul":{"id":"3502112002","nama":"Bulu Kidul"},"dadapan":{"id":"3502112009","nama":"Dadapan"},"jalen":{"id":"3502112014","nama":"Jalen"},"karangan":{"id":"3502112011","nama":"Karangan"},"karangmojo":{"id":"3502112015","nama":"Karangmojo"},"karangpatihan":{"id":"3502112005","nama":"Karangpatihan"},"muneng":{"id":"3502112019","nama":"Muneng"},"ngampel":{"id":"3502112020","nama":"Ngampel"},"ngendut":{"id":"3502112004","nama":"Ngendut"},"ngraket":{"id":"3502112008","nama":"Ngraket"},"ngumpul":{"id":"3502112007","nama":"Ngumpul"},"pandak":{"id":"3502112001","nama":"Pandak"},"purworejo":{"id":"3502112017","nama":"Purworejo"},"sedarat":{"id":"3502112016","nama":"Sedarat"},"singkil":{"id":"3502112010","nama":"Singkil"},"sumberejo":{"id":"3502112006","nama":"Sumberejo"},"tatung":{"id":"3502112018","nama":"Tatung"}}},"bungkal":{"id":"350203","nama":"Bungkal","villages":{"bancar":{"id":"3502032012","nama":"Bancar"},"bedikulon":{"id":"3502032019","nama":"Bedikulon"},"bediwetan":{"id":"3502032018","nama":"Bediwetan"},"bekare":{"id":"3502032003","nama":"Bekare"},"belang":{"id":"3502032008","nama":"Belang"},"bungkal":{"id":"3502032009","nama":"Bungkal"},"bungu":{"id":"3502032014","nama":"Bungu"},"kalisat":{"id":"3502032005","nama":"Kalisat"},"ketonggo":{"id":"3502032010","nama":"Ketonggo"},"koripan":{"id":"3502032002","nama":"Koripan"},"kunti":{"id":"3502032011","nama":"Kunti"},"kupuk":{"id":"3502032015","nama":"Kupuk"},"kwajon":{"id":"3502032017","nama":"Kwajon"},"munggu":{"id":"3502032006","nama":"Munggu"},"nambak":{"id":"3502032004","nama":"Nambak"},"padas":{"id":"3502032013","nama":"Padas"},"pager":{"id":"3502032007","nama":"Pager"},"pelem":{"id":"3502032001","nama":"Pelem"},"sambilawang":{"id":"3502032016","nama":"Sambilawang"}}},"jambon":{"id":"350220","nama":"Jambon","villages":{"blembem":{"id":"3502202009","nama":"Blembem"},"bringinan":{"id":"3502202004","nama":"Bringinan"},"bulu lor":{"id":"3502202007","nama":"Bulu Lor"},"jambon":{"id":"3502202008","nama":"Jambon"},"jonggol":{"id":"3502202001","nama":"Jonggol"},"karanglo kidul":{"id":"3502202005","nama":"Karanglo Kidul"},"krebet":{"id":"3502202006","nama":"Krebet"},"menang":{"id":"3502202011","nama":"Menang"},"poko":{"id":"3502202002","nama":"Poko"},"pulosari":{"id":"3502202010","nama":"Pulosari"},"sendang":{"id":"3502202003","nama":"Sendang"},"sidoharjo":{"id":"3502202013","nama":"Sidoharjo"},"srandil":{"id":"3502202012","nama":"Srandil"}}},"jenangan":{"id":"350218","nama":"Jenangan","villages":{"jenangan":{"id":"3502182006","nama":"Jenangan"},"jimbe":{"id":"3502182007","nama":"Jimbe"},"kemiri":{"id":"3502182017","nama":"Kemiri"},"mrican":{"id":"3502182003","nama":"Mrican"},"nglayang":{"id":"3502182005","nama":"Nglayang"},"ngrupit":{"id":"3502182008","nama":"Ngrupit"},"panjeng":{"id":"3502182011","nama":"Panjeng"},"paringan":{"id":"3502182015","nama":"Paringan"},"pintu":{"id":"3502182009","nama":"Pintu"},"plalangan":{"id":"3502182004","nama":"Plalangan"},"sedah":{"id":"3502182010","nama":"Sedah"},"semanding":{"id":"3502182013","nama":"Semanding"},"setono":{"id":"3502181001","nama":"Setono"},"singosaren":{"id":"3502181002","nama":"Singosaren"},"sraten":{"id":"3502182012","nama":"Sraten"},"tanjungsari":{"id":"3502182014","nama":"Tanjungsari"},"wates":{"id":"3502182016","nama":"Wates"}}},"jetis":{"id":"350209","nama":"Jetis","villages":{"coper":{"id":"3502092006","nama":"Coper"},"jetis":{"id":"3502092009","nama":"Jetis"},"josari":{"id":"3502092012","nama":"Josari"},"karanggebang":{"id":"3502092008","nama":"Karanggebang"},"kradenan":{"id":"3502092004","nama":"Kradenan"},"kutu kulon":{"id":"3502092002","nama":"Kutu Kulon"},"kutu wetan":{"id":"3502092003","nama":"Kutu Wetan"},"mojomati":{"id":"3502092005","nama":"Mojomati"},"mojorejo":{"id":"3502092007","nama":"Mojorejo"},"ngasinan":{"id":"3502092001","nama":"Ngasinan"},"tegalsari":{"id":"3502092010","nama":"Tegalsari"},"turi":{"id":"3502092013","nama":"Turi"},"winong":{"id":"3502092014","nama":"Winong"},"wonoketro":{"id":"3502092011","nama":"Wonoketro"}}},"kauman":{"id":"350212","nama":"Kauman","villages":{"bringin":{"id":"3502122006","nama":"Bringin"},"carat":{"id":"3502122015","nama":"Carat"},"ciluk":{"id":"3502122009","nama":"Ciluk"},"gabel":{"id":"3502122008","nama":"Gabel"},"kauman":{"id":"3502122016","nama":"Kauman"},"maron":{"id":"3502122012","nama":"Maron"},"nglarangan":{"id":"3502122005","nama":"Nglarangan"},"ngrandu":{"id":"3502122004","nama":"Ngrandu"},"nongkodono":{"id":"3502122002","nama":"Nongkodono"},"pengkol":{"id":"3502122007","nama":"Pengkol"},"plosojenar":{"id":"3502122014","nama":"Plosojenar"},"semanding":{"id":"3502122010","nama":"Semanding"},"somoroto":{"id":"3502122013","nama":"Somoroto"},"sukosari":{"id":"3502122003","nama":"Sukosari"},"tegalombo":{"id":"3502122001","nama":"Tegalombo"},"tosanan":{"id":"3502122011","nama":"Tosanan"}}},"mlarak":{"id":"350208","nama":"Mlarak","villages":{"bajang":{"id":"3502082011","nama":"Bajang"},"candi":{"id":"3502082002","nama":"Candi"},"gandu":{"id":"3502082009","nama":"Gandu"},"gontor":{"id":"3502082008","nama":"Gontor"},"jabung":{"id":"3502082010","nama":"Jabung"},"joresan":{"id":"3502082006","nama":"Joresan"},"kaponan":{"id":"3502082015","nama":"Kaponan"},"mlarak":{"id":"3502082012","nama":"Mlarak"},"nglumpang":{"id":"3502082007","nama":"Nglumpang"},"ngrukem":{"id":"3502082004","nama":"Ngrukem"},"serangan":{"id":"3502082013","nama":"Serangan"},"siwalan":{"id":"3502082005","nama":"Siwalan"},"suren":{"id":"3502082014","nama":"Suren"},"totokan":{"id":"3502082003","nama":"Totokan"},"tugu":{"id":"3502082001","nama":"Tugu"}}},"ngebel":{"id":"350219","nama":"Ngebel","villages":{"gondowido":{"id":"3502192005","nama":"Gondowido"},"ngebel":{"id":"3502192007","nama":"Ngebel"},"ngrogung":{"id":"3502192001","nama":"Ngrogung"},"pupus":{"id":"3502192006","nama":"Pupus"},"sahang":{"id":"3502192002","nama":"Sahang"},"sempu":{"id":"3502192008","nama":"Sempu"},"talun":{"id":"3502192004","nama":"Talun"},"wagir lor":{"id":"3502192003","nama":"Wagir Lor"}}},"ngrayun":{"id":"350202","nama":"Ngrayun","villages":{"baosan kidul":{"id":"3502022003","nama":"Baosan Kidul"},"baosan lor":{"id":"3502022006","nama":"Baosan Lor"},"binade":{"id":"3502022004","nama":"Binade"},"cepoko":{"id":"3502022010","nama":"Cepoko"},"gedangan":{"id":"3502022011","nama":"Gedangan"},"mrayan":{"id":"3502022005","nama":"Mrayan"},"ngrayun":{"id":"3502022009","nama":"Ngrayun"},"selur":{"id":"3502022008","nama":"Selur"},"sendang":{"id":"3502022002","nama":"Sendang"},"temon":{"id":"3502022007","nama":"Temon"},"wonodadi":{"id":"3502022001","nama":"Wonodadi"}}},"ponorogo":{"id":"350217","nama":"Ponorogo","villages":{"bangunsari":{"id":"3502171008","nama":"Bangunsari"},"banyudono":{"id":"3502171014","nama":"Banyudono"},"beduri":{"id":"3502171019","nama":"Beduri"},"brotonegaran":{"id":"3502171002","nama":"Brotonegaran"},"cokromenggalan":{"id":"3502171016","nama":"Cokromenggalan"},"jingglong":{"id":"3502171018","nama":"Jingglong"},"kauman":{"id":"3502171010","nama":"Kauman"},"keniten":{"id":"3502171017","nama":"Keniten"},"kepatihan":{"id":"3502171004","nama":"Kepatihan"},"mangkujayan":{"id":"3502171013","nama":"Mangkujayan"},"nologaten":{"id":"3502171015","nama":"Nologaten"},"paju":{"id":"3502171001","nama":"Paju"},"pakunden":{"id":"3502171003","nama":"Pakunden"},"pinggirsari":{"id":"3502171012","nama":"Pinggirsari"},"purbosuman":{"id":"3502171006","nama":"Purbosuman"},"surodikraman":{"id":"3502171005","nama":"Surodikraman"},"tamanarum":{"id":"3502171009","nama":"Tamanarum"},"tambakbayan":{"id":"3502171011","nama":"Tambakbayan"},"tonatan":{"id":"3502171007","nama":"Tonatan"}}},"pudak":{"id":"350221","nama":"Pudak","villages":{"banjarejo":{"id":"3502212005","nama":"Banjarejo"},"bareng":{"id":"3502212001","nama":"Bareng"},"krisik":{"id":"3502212003","nama":"Krisik"},"pudak kulon":{"id":"3502212006","nama":"Pudak Kulon"},"pudak wetan":{"id":"3502212004","nama":"Pudak Wetan"},"tambang":{"id":"3502212002","nama":"Tambang"}}},"pulung":{"id":"350207","nama":"Pulung","villages":{"banaran":{"id":"3502072018","nama":"Banaran"},"bedrug":{"id":"3502072003","nama":"Bedrug"},"bekiring":{"id":"3502072017","nama":"Bekiring"},"karangpatihan":{"id":"3502072001","nama":"Karangpatihan"},"kesugihan":{"id":"3502072013","nama":"Kesugihan"},"munggung":{"id":"3502072016","nama":"Munggung"},"patik":{"id":"3502072006","nama":"Patik"},"plunturan":{"id":"3502072011","nama":"Plunturan"},"pomahan":{"id":"3502072012","nama":"Pomahan"},"pulung":{"id":"3502072007","nama":"Pulung"},"pulung merdiko":{"id":"3502072008","nama":"Pulung Merdiko"},"serag":{"id":"3502072014","nama":"Serag"},"sidoharjo":{"id":"3502072009","nama":"Sidoharjo"},"singgahan":{"id":"3502072005","nama":"Singgahan"},"tegalrejo":{"id":"3502072002","nama":"Tegalrejo"},"wagir kidul":{"id":"3502072004","nama":"Wagir Kidul"},"wayang":{"id":"3502072015","nama":"Wayang"},"wotan":{"id":"3502072010","nama":"Wotan"}}},"sambit":{"id":"350204","nama":"Sambit","villages":{"bancangan":{"id":"3502042007","nama":"Bancangan"},"bangsalan":{"id":"3502042014","nama":"Bangsalan"},"bedingin":{"id":"3502042006","nama":"Bedingin"},"besuki":{"id":"3502042012","nama":"Besuki"},"bulu":{"id":"3502042010","nama":"Bulu"},"campurejo":{"id":"3502042008","nama":"Campurejo"},"campursari":{"id":"3502042009","nama":"Campursari"},"gajah":{"id":"3502042001","nama":"Gajah"},"jrakah":{"id":"3502042016","nama":"Jrakah"},"kemuning":{"id":"3502042015","nama":"Kemuning"},"maguwan":{"id":"3502042004","nama":"Maguwan"},"ngadisanan":{"id":"3502042003","nama":"Ngadisanan"},"nglewan":{"id":"3502042005","nama":"Nglewan"},"sambit":{"id":"3502042011","nama":"Sambit"},"wilangan":{"id":"3502042013","nama":"Wilangan"},"wringinanom":{"id":"3502042002","nama":"Wringinanom"}}},"sampung":{"id":"350214","nama":"Sampung","villages":{"carangrejo":{"id":"3502142006","nama":"Carangrejo"},"gelangkulon":{"id":"3502142010","nama":"Gelangkulon"},"glinggang":{"id":"3502142009","nama":"Glinggang"},"jenangan":{"id":"3502142003","nama":"Jenangan"},"karangwaluh":{"id":"3502142011","nama":"Karangwaluh"},"kunti":{"id":"3502142008","nama":"Kunti"},"nglurup":{"id":"3502142004","nama":"Nglurup"},"pagerukir":{"id":"3502142007","nama":"Pagerukir"},"pohijo":{"id":"3502142002","nama":"Pohijo"},"ringinputih":{"id":"3502142012","nama":"Ringinputih"},"sampung":{"id":"3502142001","nama":"Sampung"},"tulung":{"id":"3502142005","nama":"Tulung"}}},"sawoo":{"id":"350205","nama":"Sawoo","villages":{"bondrang":{"id":"3502052013","nama":"Bondrang"},"grogol":{"id":"3502052010","nama":"Grogol"},"ketro":{"id":"3502052011","nama":"Ketro"},"kori":{"id":"3502052012","nama":"Kori"},"ngindeng":{"id":"3502052014","nama":"Ngindeng"},"pangkal":{"id":"3502052002","nama":"Pangkal"},"prayungan":{"id":"3502052008","nama":"Prayungan"},"sawoo":{"id":"3502052007","nama":"Sawoo"},"sriti":{"id":"3502052005","nama":"Sriti"},"temon":{"id":"3502052006","nama":"Temon"},"tempuran":{"id":"3502052004","nama":"Tempuran"},"tugurejo":{"id":"3502052009","nama":"Tugurejo"},"tumpak pelem":{"id":"3502052003","nama":"Tumpak Pelem"},"tumpuk":{"id":"3502052001","nama":"Tumpuk"}}},"siman":{"id":"350210","nama":"Siman","villages":{"beton":{"id":"3502102004","nama":"Beton"},"brahu":{"id":"3502102006","nama":"Brahu"},"demangan":{"id":"3502102001","nama":"Demangan"},"jarak":{"id":"3502102009","nama":"Jarak"},"kepuhrubuh":{"id":"3502102007","nama":"Kepuhrubuh"},"madusari":{"id":"3502102003","nama":"Madusari"},"mangunsuman":{"id":"3502101018","nama":"Mangunsuman"},"manuk":{"id":"3502102012","nama":"Manuk"},"ngabar":{"id":"3502102002","nama":"Ngabar"},"patihan kidul":{"id":"3502102014","nama":"Patihan Kidul"},"pijeran":{"id":"3502102011","nama":"Pijeran"},"ronosentanan":{"id":"3502102015","nama":"Ronosentanan"},"ronowijayan":{"id":"3502101017","nama":"Ronowijayan"},"sawuh":{"id":"3502102008","nama":"Sawuh"},"sekaran":{"id":"3502102005","nama":"Sekaran"},"siman":{"id":"3502102013","nama":"Siman"},"tajug":{"id":"3502102016","nama":"Tajug"},"tranjang":{"id":"3502102010","nama":"Tranjang"}}},"slahung":{"id":"350201","nama":"Slahung","villages":{"broto":{"id":"3502012005","nama":"Broto"},"caluk":{"id":"3502012004","nama":"Caluk"},"crabak":{"id":"3502012017","nama":"Crabak"},"duri":{"id":"3502012010","nama":"Duri"},"galak":{"id":"3502012014","nama":"Galak"},"gombang":{"id":"3502012021","nama":"Gombang"},"gundik":{"id":"3502012019","nama":"Gundik"},"janti":{"id":"3502012022","nama":"Janti"},"jebeng":{"id":"3502012013","nama":"Jebeng"},"kambeng":{"id":"3502012007","nama":"Kambeng"},"menggare":{"id":"3502012006","nama":"Menggare"},"mojopitu":{"id":"3502012018","nama":"Mojopitu"},"nailan":{"id":"3502012020","nama":"Nailan"},"ngilo-ilo":{"id":"3502012009","nama":"Ngilo-Ilo"},"ngloning":{"id":"3502012011","nama":"Ngloning"},"plancungan":{"id":"3502012012","nama":"Plancungan"},"senepo":{"id":"3502012002","nama":"Senepo"},"simo":{"id":"3502012016","nama":"Simo"},"slahung":{"id":"3502012003","nama":"Slahung"},"truneng":{"id":"3502012015","nama":"Truneng"},"tugurejo":{"id":"3502012001","nama":"Tugurejo"},"wates":{"id":"3502012008","nama":"Wates"}}},"sooko":{"id":"350206","nama":"Sooko","villages":{"bedoho":{"id":"3502062005","nama":"Bedoho"},"jurug":{"id":"3502062006","nama":"Jurug"},"klepu":{"id":"3502062002","nama":"Klepu"},"ngadirojo":{"id":"3502062001","nama":"Ngadirojo"},"sooko":{"id":"3502062004","nama":"Sooko"},"suru":{"id":"3502062003","nama":"Suru"}}},"sukorejo":{"id":"350215","nama":"Sukorejo","villages":{"bangunrejo":{"id":"3502152002","nama":"Bangunrejo"},"gandukepuh":{"id":"3502152015","nama":"Gandukepuh"},"gegeran":{"id":"3502152014","nama":"Gegeran"},"gelang lor":{"id":"3502152006","nama":"Gelang Lor"},"golan":{"id":"3502152010","nama":"Golan"},"kalimalang":{"id":"3502152016","nama":"Kalimalang"},"karanglo lor":{"id":"3502152013","nama":"Karanglo Lor"},"kedungbanteng":{"id":"3502152012","nama":"Kedungbanteng"},"kranggan":{"id":"3502152009","nama":"Kranggan"},"lengkong":{"id":"3502152005","nama":"Lengkong"},"morosari":{"id":"3502152018","nama":"Morosari"},"nambangrejo":{"id":"3502152004","nama":"Nambangrejo"},"nampan":{"id":"3502152008","nama":"Nampan"},"prajegan":{"id":"3502152011","nama":"Prajegan"},"serangan":{"id":"3502152007","nama":"Serangan"},"sidorejo":{"id":"3502152003","nama":"Sidorejo"},"sragi":{"id":"3502152017","nama":"Sragi"},"sukorejo":{"id":"3502152001","nama":"Sukorejo"}}}}', true); return $m; }
    function sim_norm($s){ return strtolower(trim(preg_replace('/\s+/', ' ', (string)$s))); }
    
    function sim_region_options($ch, $jenis, $kode){
        $url = SIMLINMAS_BASE . '/region/' . $jenis . '/' . $kode;
        [$st, $html, $err] = curl_req($ch, 'GET', $url, null, [
            'Referer: ' . SIMLINMAS_BASE . '/linmas/tambah',
            'Accept: */*',
            'X-Requested-With: XMLHttpRequest',
        ]);
        if ($err || is_cloudflare_blocked($html, $st)) return [];
        
        preg_match_all('/<option\s+value="([^"]*)"[^>]*>([^<]*)<\/option>/i', $html, $m, PREG_SET_ORDER);
        $out = [];
        foreach ($m as $r){ 
            $v = trim($r[1]); 
            $t = trim($r[2]); 
            if ($v !== '' && $v !== '0') $out[] = ['id' => $v, 'nama' => $t]; 
        }
        return $out;
    }
    
    function sim_resolve_wilayah($ch, &$fields){
        $kecNama = trim((string)($fields['kecamatan_nama'] ?? $fields['_kecamatan_nama'] ?? ''));
        $kelNama = trim((string)($fields['kelurahan_nama'] ?? $fields['_kelurahan_nama'] ?? ''));
        unset($fields['kecamatan_nama'], $fields['kelurahan_nama'], $fields['_kecamatan_nama'], $fields['_kelurahan_nama']);
        $kab = (string)($fields['kabkota_id'] ?? '3502');
        $map = wilayah_map();
        if (empty($fields['kecamatan_id']) && $kecNama !== '') {
            $k = $map[sim_norm($kecNama)] ?? null;
            if ($k) { $fields['kecamatan_id'] = $k['id']; }
        }
        if (empty($fields['kelurahan_id']) && !empty($fields['kecamatan_id']) && $kelNama !== '') {
            $kecId = (string)$fields['kecamatan_id'];
            $kecEntry = null;
            foreach ($map as $mk) { if ((string)$mk['id'] === $kecId) { $kecEntry = $mk; break; } }
            if ($kecEntry) {
                $v = $kecEntry['villages'][sim_norm($kelNama)] ?? null;
                if ($v) { $fields['kelurahan_id'] = $v['id']; }
            }
        }
        // fallback live
        if (empty($fields['kecamatan_id']) && $kecNama !== '') {
            $opts = sim_region_options($ch, 'districts', $kab);
            foreach ($opts as $o){ if (sim_norm($o['nama']) === sim_norm($kecNama)) { $fields['kecamatan_id'] = $o['id']; break; } }
        }
        if (empty($fields['kelurahan_id']) && !empty($fields['kecamatan_id']) && $kelNama !== '') {
            $opts = sim_region_options($ch, 'villages', (string)$fields['kecamatan_id']);
            foreach ($opts as $o){ if (sim_norm($o['nama']) === sim_norm($kelNama)) { $fields['kelurahan_id'] = $o['id']; break; } }
        }
    }

    if ($action === 'resolve_wilayah') {
        $fields = $_POST;
        sim_resolve_wilayah($ch, $fields);
        out(['ok' => true, 'action' => 'resolve_wilayah', 'kecamatan_id' => $fields['kecamatan_id'] ?? '', 'kelurahan_id' => $fields['kelurahan_id'] ?? '']);
    }

    if ($action === 'insert') {
        $tambahUrl = SIMLINMAS_BASE . '/linmas/tambah';
        [$s3, $htmlTambah, $e3] = curl_req($ch, 'GET', $tambahUrl, null, get_default_headers(SIMLINMAS_BASE . '/linmas'));
        if ($e3) out(['ok' => false, 'message' => 'Gagal buka form tambah: ' . $e3], 502);
        if (is_cloudflare_blocked($htmlTambah, $s3)) out(['ok' => false, 'message' => 'Diblokir Cloudflare saat membuka form tambah.'], 502);
        
        $csrf = get_csrf($htmlTambah);
        if ($csrf === '') out(['ok' => false, 'message' => 'CSRF tambah tidak ditemukan'], 502);

        $fields = $_POST;
        unset($fields['token'], $fields['action']);
        $fields['_token']     = $csrf;
        $fields['provinsi_id'] = $fields['provinsi_id'] ?? '35';
        $fields['kabkota_id']  = $fields['kabkota_id'] ?? '3502';
        sim_resolve_wilayah($ch, $fields);

        foreach (['foto_profil', 'sk_file'] as $f) {
            if (isset($_FILES[$f]) && is_array($_FILES[$f]) && ($_FILES[$f]['error'] ?? UPLOAD_ERR_NO_FILE) === UPLOAD_ERR_OK) {
                $fields[$f] = new CURLFile($_FILES[$f]['tmp_name'], $_FILES[$f]['type'] ?: 'application/octet-stream', $_FILES[$f]['name'] ?: $f);
            }
        }

        $insertUrl = SIMLINMAS_BASE . '/linmas/insert';
        [$s4, $body4, $e4] = curl_req($ch, 'POST', $insertUrl, $fields, [
            'X-CSRF-TOKEN: ' . $csrf,
            'X-Requested-With: XMLHttpRequest',
            'Referer: ' . $tambahUrl,
            'Origin: ' . SIMLINMAS_BASE,
            'Accept: application/json, text/javascript, */*; q=0.01',
            'Sec-Ch-Ua: "Chromium";v="122", "Not(A:Brand";v="8", "Google Chrome";v="122"',
            'Sec-Ch-Ua-Mobile: ?0',
            'Sec-Ch-Ua-Platform: "Windows"',
        ]);
        
        if ($e4) out(['ok' => false, 'action' => 'insert', 'message' => 'Gagal kirim ke SIMLINMAS: ' . $e4], 502);
        if (is_cloudflare_blocked($body4, $s4)) out(['ok' => false, 'action' => 'insert', 'message' => 'Aksi insert diblokir oleh Cloudflare WAF/Bot Protection.'], 502);

        $json = json_decode($body4, true);
        if (is_array($json)) {
            $ok = (string)($json['code'] ?? '') === '200';
            out(['ok' => $ok, 'action' => 'insert', 'code' => $json['code'] ?? null, 'message' => $json['message'] ?? $body4, 'raw' => $json], 200);
        }
        
        out(['ok' => false, 'action' => 'insert', 'message' => 'Respons SIMLINMAS bukan JSON yang valid', 'raw' => substr($body4, 0, 500)], 200);
    }

    if ($action === 'update') {
        $idEdit = trim((string)($_POST['id_edit'] ?? $_POST['id'] ?? ''));
        if ($idEdit === '') out(['ok'=>false,'action'=>'update','message'=>'id_edit wajib'], 400);
        $editUrl = SIMLINMAS_BASE . '/linmas/edit/' . rawurlencode($idEdit);
        [$s3, $htmlEdit, $e3] = curl_req($ch, 'GET', $editUrl, null, get_default_headers(SIMLINMAS_BASE . '/linmas'));
        if ($e3) out(['ok'=>false,'action'=>'update','message'=>'Gagal buka form edit: '.$e3], 502);
        if (is_cloudflare_blocked($htmlEdit, $s3)) out(['ok'=>false,'action'=>'update','message'=>'Diblokir Cloudflare saat membuka form edit.'], 502);
        $csrf = get_csrf($htmlEdit);
        if ($csrf === '') out(['ok'=>false,'action'=>'update','message'=>'CSRF edit tidak ditemukan'], 502);

        $fields = $_POST;
        unset($fields['token'], $fields['action']);
        $fields['_token'] = $csrf;
        $fields['id_edit'] = $idEdit;
        $fields['provinsi_id'] = $fields['provinsi_id'] ?? '35';
        $fields['kabkota_id']  = $fields['kabkota_id'] ?? '3502';
        sim_resolve_wilayah($ch, $fields);

        foreach (['foto_profil','sk_file'] as $f) {
            if (isset($_FILES[$f]) && is_array($_FILES[$f]) && ($_FILES[$f]['error'] ?? UPLOAD_ERR_NO_FILE) === UPLOAD_ERR_OK) {
                $fields[$f] = new CURLFile($_FILES[$f]['tmp_name'], $_FILES[$f]['type'] ?: 'application/octet-stream', $_FILES[$f]['name'] ?: $f);
            }
        }

        $updateUrl = SIMLINMAS_BASE . '/linmas/update';
        [$s4, $body4, $e4] = curl_req($ch, 'POST', $updateUrl, $fields, [
            'X-CSRF-TOKEN: ' . $csrf,
            'X-Requested-With: XMLHttpRequest',
            'Referer: ' . $editUrl,
            'Origin: ' . SIMLINMAS_BASE,
            'Accept: application/json, text/javascript, */*; q=0.01',
        ]);
        if ($e4) out(['ok'=>false,'action'=>'update','message'=>'Gagal kirim update ke SIMLINMAS: '.$e4], 502);
        if (is_cloudflare_blocked($body4, $s4)) out(['ok'=>false,'action'=>'update','message'=>'Aksi update diblokir Cloudflare.'], 502);
        $json = json_decode($body4, true);
        if (is_array($json)) {
            $ok = (string)($json['code'] ?? '') === '200';
            out(['ok'=>$ok,'action'=>'update','code'=>$json['code'] ?? null,'message'=>$json['message'] ?? $body4,'raw'=>$json], 200);
        }
        out(['ok'=>false,'action'=>'update','message'=>'Respons SIMLINMAS bukan JSON valid','raw'=>substr($body4,0,500)], 200);
    }

    if ($action === 'kta_members') {
        $kecFilter = sim_norm((string)($_POST['kecamatan'] ?? ''));
        $kelFilter = sim_norm((string)($_POST['kelurahan'] ?? $_POST['desa'] ?? ''));
        $kecIdFilter = trim((string)($_POST['kecamatan_id'] ?? ''));
        $kelIdFilter = trim((string)($_POST['kelurahan_id'] ?? ''));
        // resolve name -> id from hardcoded map
        $map = wilayah_map();
        if ($kecIdFilter === '' && $kecFilter !== '' && isset($map[$kecFilter])) $kecIdFilter = (string)$map[$kecFilter]['id'];
        if ($kelIdFilter === '' && $kecFilter !== '' && $kelFilter !== '' && isset($map[$kecFilter]['villages'][$kelFilter])) $kelIdFilter = (string)$map[$kecFilter]['villages'][$kelFilter]['id'];
        $ajaxHeaders = [
            'Accept: application/json, text/javascript, */*; q=0.01',
            'X-Requested-With: XMLHttpRequest',
            'Referer: ' . SIMLINMAS_BASE . '/linmas',
        ];
        [$st0,$body0,$err0] = curl_req($ch,'GET', SIMLINMAS_BASE.'/linmas?start=0&length=1', null, $ajaxHeaders);
        $total = 0; $j0 = json_decode($body0, true);
        if (is_array($j0)) $total = (int)($j0['recordsFiltered'] ?? $j0['recordsTotal'] ?? 0);
        $rows = []; $start = 0;
        while ($start <= ($total > 0 ? $total + 400 : 5000)) {
            [$st,$body,$err] = curl_req($ch,'GET', SIMLINMAS_BASE.'/linmas?start='.$start.'&length=200', null, $ajaxHeaders);
            if ($err) break;
            if (is_cloudflare_blocked($body,$st)) break;
            $j = json_decode($body, true);
            $data = is_array($j) ? ($j['data'] ?? []) : [];
            if (!is_array($data) || count($data)==0) break;
            foreach ($data as $r) {
                $rkid = trim((string)($r['kecamatan_id'] ?? ''));
                $rlid = trim((string)($r['kelurahan_id'] ?? ''));
                $kecName = is_array($r['kecamatan']??null) ? (string)($r['kecamatan']['name']??'') : (string)($r['kecamatan'] ?? '');
                $kelName = is_array($r['desakel']??null) ? (string)($r['desakel']['name']??'') : (string)($r['desakel'] ?? $r['kelurahan'] ?? '');
                if ($kecIdFilter !== '' && $rkid !== $kecIdFilter) continue;
                if ($kelIdFilter !== '' && $rlid !== $kelIdFilter) continue;
                if ($kecIdFilter === '' && $kecFilter !== '' && sim_norm($kecName) !== $kecFilter) continue;
                if ($kelIdFilter === '' && $kelFilter !== '' && sim_norm($kelName) !== $kelFilter) continue;
                $rows[] = [
                    'id' => (string)($r['id'] ?? ''),
                    'nik' => (string)($r['nik'] ?? ''),
                    'nama_lengkap' => (string)($r['nama_lengkap'] ?? ''),
                    'nomor_registrasi' => (string)($r['nomor_registrasi'] ?? ''),
                    'kecamatan' => $kecName,
                    'kelurahan' => $kelName,
                    'jabatan' => (string)($r['jabatan'] ?? ''),
                    'status' => (string)($r['status'] ?? ''),
                ];
            }
            $start += 200;
            if ($total===0 && is_array($j)) $total = (int)($j['recordsFiltered'] ?? $j['recordsTotal'] ?? 0);
        }
        out(['ok'=>true,'action'=>'kta_members','count'=>count($rows),'members'=>$rows]);
    }

    if ($action === 'kta_data') {
        $id = trim((string)($_POST['id'] ?? ''));
        if ($id === '') out(['ok'=>false,'action'=>'kta_data','message'=>'id wajib'], 400);
        $detailUrl = SIMLINMAS_BASE . '/linmas/detail/' . rawurlencode($id);
        [$st,$body,$err] = curl_req($ch,'GET',$detailUrl, null, [
            'Accept: application/json, text/plain, */*',
            'X-Requested-With: XMLHttpRequest',
            'Referer: ' . SIMLINMAS_BASE . '/linmas',
        ]);
        if ($err) out(['ok'=>false,'action'=>'kta_data','message'=>'Gagal ambil detail: '.$err], 502);
        $detail = json_decode($body, true);
        if (!is_array($detail)) out(['ok'=>false,'action'=>'kta_data','message'=>'Detail bukan JSON','raw'=>substr($body,0,300)], 502);
        $u = $detail['user'] ?? [];
        $fetchB64 = function($url) use ($ch) {
            $url = trim((string)$url);
            if ($url === '' || $url === '-' || $url === 'null') return '';
            if (strpos($url, '/') === 0) $url = SIMLINMAS_BASE . $url;
            if (strpos($url, 'http') !== 0) $url = SIMLINMAS_BASE . '/' . $url;
            [$s,$b,$e] = curl_req($ch,'GET',$url, null, ['Accept: image/*,*/*;q=0.8', 'Referer: '.SIMLINMAS_BASE.'/linmas']);
            if ($e || $s < 200 || $s >= 300 || $b === '') return '';
            $ct = 'image/png';
            if (strncmp($b, "\xFF\xD8\xFF", 3)===0) $ct='image/jpeg';
            elseif (strncmp($b, "\x89PNG\r\n\x1a\n", 8)===0) $ct='image/png';
            elseif (strpos($b, '<svg') === 0 || strpos($b, '<?xml') === 0) $ct='image/svg+xml';
            return 'data:'.$ct.';base64,'.base64_encode($b);
        };
        $kab = $detail['kabkota'] ?? [];
        $assets = [
            'foto' => $fetchB64($u['foto_profil'] ?? ''),
            'logo_kem' => $fetchB64('/logo-kemendagri.png'),
            'logo_kab' => $fetchB64($kab['logo'] ?? ''),
            'wm' => $fetchB64('/logo.png'),
        ];
        out(['ok'=>true,'action'=>'kta_data','id'=>$id,'detail'=>$detail,'assets'=>$assets]);
    }

    out(['ok' => false, 'message' => 'Action tidak dikenal: ' . $action], 400);
} finally {
    curl_close($ch);
    @unlink($cookie);
}
