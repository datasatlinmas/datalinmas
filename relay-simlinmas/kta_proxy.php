<?php
/**
 * Proxy Download KTA - server-side token + verifikasi No HP admin.
 * Action diizinkan: kta_members, kta_data.
 * Wajib kirim: kecamatan, kelurahan/desa (opsional untuk level kecamatan), hp_admin.
 */
declare(strict_types=1);
header('Content-Type: application/json; charset=utf-8');
header('X-Content-Type-Options: nosniff');

$origin = $_SERVER['HTTP_ORIGIN'] ?? '';
if ($origin !== '') {
    $host = parse_url($origin, PHP_URL_HOST);
    $self = $_SERVER['HTTP_HOST'] ?? '';
    if ($host === $self) {
        header('Access-Control-Allow-Origin: ' . $origin);
        header('Vary: Origin');
    }
}
header('Access-Control-Allow-Methods: POST, OPTIONS');
header('Access-Control-Allow-Headers: Content-Type');
if ($_SERVER['REQUEST_METHOD'] === 'OPTIONS') { http_response_code(204); exit; }
if ($_SERVER['REQUEST_METHOD'] !== 'POST') {
    http_response_code(405); echo json_encode(['ok'=>false,'message'=>'Method harus POST']); exit;
}

const RELAY_URL = 'https://satpolpp.ponorogo.go.id/satlinmas/relay/simlinmas_relay.php';
const RELAY_TOKEN = 'jaUIJ21PG8qJRKeegiyZEFywBApjcMeW3LWVMJ61iI0';
const ADMIN_API = 'https://script.google.com/macros/s/AKfycbwGke0-djhNYx-Gk4__sLvjs9sRovvKpHw56hf_m9PHc0Nd_f31UXJ3X4bBhD9VcOZZsA/exec?api=admin&aksi=list';

function jout($data, $code=200){ http_response_code($code); echo json_encode($data, JSON_UNESCAPED_UNICODE|JSON_UNESCAPED_SLASHES); exit; }
function norm_hp($s){ $s = preg_replace('/\D+/', '', (string)$s); if (strpos($s,'62')===0) $s = '0'.substr($s,2); elseif (strpos($s,'8')===0) $s = '0'.$s; return $s; }
function norm_txt($s){ return strtolower(trim(preg_replace('/\s+/', ' ', (string)$s))); }

function client_ip(){
    $ip = (string)($_SERVER['REMOTE_ADDR'] ?? 'unknown');
    return preg_replace('/[^a-fA-F0-9:\.]/', '_', $ip);
}
function rl_path(){ return sys_get_temp_dir() . '/kta_rl_' . hash('sha256', client_ip()) . '.json'; }
function rl_load(){
    $f = rl_path();
    $d = is_file($f) ? json_decode((string)@file_get_contents($f), true) : null;
    if (!is_array($d)) $d = ['fails'=>0,'blocked_until'=>0];
    return $d;
}
function rl_save($d){ @file_put_contents(rl_path(), json_encode($d), LOCK_EX); }
function rl_check(){
    $d = rl_load(); $now = time();
    if (!empty($d['blocked_until']) && $d['blocked_until'] > $now) {
        $wait = $d['blocked_until'] - $now;
        $menit = (int)ceil($wait/60);
        jout(['ok'=>false,'message'=>'Terlalu banyak percobaan No. HP salah. Coba lagi dalam '.$menit.' menit ('.$wait.' detik).','retry_after'=>$wait], 429);
    }
    if (!empty($d['blocked_until']) && $d['blocked_until'] <= $now) { $d['fails']=0; $d['blocked_until']=0; rl_save($d); }
}
function rl_fail(){
    $d = rl_load(); $d['fails'] = (int)($d['fails'] ?? 0) + 1;
    if ($d['fails'] >= 5) { $d['blocked_until'] = time() + 5*60; $d['fails']=0; rl_save($d); jout(['ok'=>false,'message'=>'Terlalu banyak percobaan No. HP salah (5x). Tunggu 5 menit sebelum mencoba lagi.','retry_after'=>300], 429); }
    rl_save($d);
    $sisa = 5 - $d['fails'];
    return $sisa;
}
function rl_clear(){ $d=['fails'=>0,'blocked_until'=>0]; rl_save($d); }

$allowed = ['kta_members', 'kta_data'];
$action = trim((string)($_POST['action'] ?? ''));
if (!in_array($action, $allowed, true)) jout(['ok'=>false,'message'=>'Action tidak diizinkan'], 400);

rl_check();

$kec = trim((string)($_POST['kecamatan'] ?? ''));
$desa = trim((string)($_POST['kelurahan'] ?? $_POST['desa'] ?? ''));
$hpInput = norm_hp($_POST['hp_admin'] ?? $_POST['no_hp'] ?? '');
if ($kec === '') jout(['ok'=>false,'message'=>'Kecamatan wajib dipilih'], 400);
if ($hpInput === '') { $sisa = rl_fail(); jout(['ok'=>false,'message'=>'No. HP admin wajib diisi untuk konfirmasi. Sisa percobaan: '.$sisa], 400); }

// Ambil data admin dari Apps Script (cache 60 detik)
$cacheFile = sys_get_temp_dir() . '/kta_admin_cache.json';
$cacheTtl = 60;
$adminJson = null;
if (is_file($cacheFile) && (time() - filemtime($cacheFile) < $cacheTtl)) {
    $adminJson = json_decode((string)@file_get_contents($cacheFile), true);
}
if (!is_array($adminJson) || !($adminJson['ok'] ?? false)) {
    $ch = curl_init(ADMIN_API);
    curl_setopt_array($ch, [
        CURLOPT_RETURNTRANSFER => true,
        CURLOPT_FOLLOWLOCATION => true,
        CURLOPT_TIMEOUT => 30,
        CURLOPT_CONNECTTIMEOUT => 15,
        CURLOPT_SSL_VERIFYPEER => true,
        CURLOPT_SSL_VERIFYHOST => 2,
        CURLOPT_HTTPHEADER => ['Accept: application/json,text/plain,*/*'],
    ]);
    $adminBody = curl_exec($ch);
    $adminErr = curl_error($ch);
    curl_close($ch);
    if ($adminBody === false || $adminErr) jout(['ok'=>false,'message'=>'Gagal memuat data admin: '.$adminErr], 502);
    $adminJson = json_decode((string)$adminBody, true);
    if (!is_array($adminJson)) {
        if (preg_match('/^[a-zA-Z0-9_]+\((.*)\)\s*;?\s*$/s', trim((string)$adminBody), $m)) $adminJson = json_decode($m[1], true);
    }
    if (is_array($adminJson) && ($adminJson['ok'] ?? false)) @file_put_contents($cacheFile, json_encode($adminJson));
}
if (!is_array($adminJson) || !($adminJson['ok'] ?? false)) jout(['ok'=>false,'message'=>'Data admin tidak valid'], 502);
$rows = is_array($adminJson['data'] ?? null) ? $adminJson['data'] : [];

$kecN = norm_txt($kec); $desaN = norm_txt($desa);
$matches = [];
foreach ($rows as $r) {
    $rk = norm_txt($r['kecamatan'] ?? '');
    $rd = norm_txt($r['desa'] ?? '');
    $rj = norm_txt($r['jenis'] ?? '');
    if ($rk !== $kecN) continue;
    if ($desaN !== '') {
        if (strpos($rj,'desa') !== false && $rd === $desaN) $matches[] = $r;
    } else {
        if (strpos($rj,'kecamatan') !== false) $matches[] = $r;
    }
}
$label = $desaN !== '' ? ('Admin Desa '.$desa.' (Kec. '.$kec.')') : ('Admin Kecamatan '.$kec);
if (count($matches) === 0) jout(['ok'=>false,'message'=>$label.' belum terdaftar. Lengkapi Data Admin terlebih dahulu.'], 403);
$hpOk = false;
$hasHp = false;
foreach ($matches as $a) {
    $hp = norm_hp($a['wa'] ?? $a['no_hp'] ?? $a['telepon'] ?? '');
    if ($hp !== '') { $hasHp = true; if ($hp === $hpInput) { $hpOk = true; break; } }
}
if (!$hasHp) jout(['ok'=>false,'message'=>$label.' belum memiliki No. HP. Lengkapi Data Admin terlebih dahulu.'], 403);
if (!$hpOk) { $sisa = rl_fail(); jout(['ok'=>false,'message'=>'No. HP tidak sesuai dengan '.$label.'. Download dibatalkan. Sisa percobaan: '.$sisa,'remaining'=>$sisa], 403); }

// Verifikasi OK
rl_clear();

// Forward ke relay
$post = $_POST;
$post['token'] = RELAY_TOKEN;
$post['action'] = $action;
unset($post['hp_admin'], $post['no_hp']);

$ch = curl_init(RELAY_URL);
curl_setopt_array($ch, [
    CURLOPT_RETURNTRANSFER => true,
    CURLOPT_POST => true,
    CURLOPT_POSTFIELDS => http_build_query($post),
    CURLOPT_TIMEOUT => 180,
    CURLOPT_CONNECTTIMEOUT => 30,
    CURLOPT_FOLLOWLOCATION => true,
    CURLOPT_SSL_VERIFYPEER => true,
    CURLOPT_SSL_VERIFYHOST => 2,
    CURLOPT_HTTPHEADER => ['Accept: application/json', 'Content-Type: application/x-www-form-urlencoded'],
]);
$body = curl_exec($ch);
$err = curl_error($ch);
$status = (int)curl_getinfo($ch, CURLINFO_HTTP_CODE);
curl_close($ch);

if ($body === false || $err) jout(['ok'=>false,'message'=>'Gagal menghubungi relay: '.$err], 502);
http_response_code($status > 0 ? $status : 200);
$json = json_decode((string)$body, true);
if (is_array($json)) echo json_encode($json, JSON_UNESCAPED_UNICODE|JSON_UNESCAPED_SLASHES);
else echo json_encode(['ok'=>false,'message'=>'Respons relay tidak valid','raw'=>substr((string)$body,0,300)], JSON_UNESCAPED_UNICODE);
