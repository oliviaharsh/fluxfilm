<?php
// api.php — FluxFilm Hostinger Proxy -> Google Apps Script

ini_set('display_errors', 0);
error_reporting(E_ALL);

// ── Secrets loaded from gitignored config.php ──
$CONFIG = require __DIR__ . '/config.php';
$API_KEY         = $CONFIG['API_KEY'];
$APPS_SCRIPT_URL = $CONFIG['APPS_SCRIPT_URL'];
$CACHE_CLEAR_KEY = $CONFIG['CACHE_CLEAR_KEY'];

$CACHE_DIR = __DIR__ . '/cache/';
$CACHEABLE = ['getBootstrap', 'getStockLevels', 'getTrendingItems'];
$CACHE_TTL = 60; // seconds

header("Content-Type: application/json");

// ── Cache clear endpoint ──
// Bookmark: https://yoursite.com/api.php?clearcache=1&key=YOUR_CACHE_CLEAR_KEY
if (isset($_GET['clearcache']) && $_GET['key'] === $CACHE_CLEAR_KEY) {
  if (is_dir($CACHE_DIR)) array_map('unlink', glob($CACHE_DIR . '*.json'));
  echo json_encode(["ok" => true, "message" => "Cache cleared"]);
  exit;
}

// ---- OPTIONS (CORS) ----
if ($_SERVER["REQUEST_METHOD"] === "OPTIONS") {
  header("Access-Control-Allow-Origin: *");
  header("Access-Control-Allow-Headers: Content-Type, X-API-KEY");
  header("Access-Control-Allow-Methods: POST, OPTIONS");
  http_response_code(200);
  echo json_encode(["ok" => true]);
  exit;
}

header("Access-Control-Allow-Origin: *");

// Read payload
$payload = file_get_contents("php://input");
if (!$payload) $payload = "{}";

$body = json_decode($payload, true);
if (!is_array($body)) $body = [];

$action = isset($body["action"]) ? strval($body["action"]) : "";

// ── Serve from cache if available ──
if (in_array($action, $CACHEABLE)) {
  $cacheFile = $CACHE_DIR . $action . '.json';
  if (file_exists($cacheFile) && (time() - filemtime($cacheFile)) < $CACHE_TTL) {
    echo file_get_contents($cacheFile);
    exit;
  }
}

// ✅ Only these actions are allowed without API key
$publicActions = ["getFaqs"];

// Read API key
$key = "";
if (function_exists("getallheaders")) {
  $h = array_change_key_case(getallheaders(), CASE_LOWER);
  $key = $h["x-api-key"] ?? "";
}
if (!$key) $key = $_SERVER["HTTP_X_API_KEY"] ?? $_SERVER["REDIRECT_HTTP_X_API_KEY"] ?? "";
if (!$key && isset($body["apiKey"])) $key = strval($body["apiKey"]);

// Enforce auth unless action is public
if (!in_array($action, $publicActions, true)) {
  if (strval($key) !== $API_KEY) {
    http_response_code(403);
    echo json_encode(["ok" => false, "message" => "Unauthorized"]);
    exit;
  }
}

// Add API key into body for GAS router
$body["apiKey"] = $API_KEY;
$payload = json_encode($body);

if (!function_exists("curl_init")) {
  http_response_code(500);
  echo json_encode(["ok" => false, "message" => "PHP cURL extension missing"]);
  exit;
}

$ch = curl_init($APPS_SCRIPT_URL);
curl_setopt_array($ch, [
  CURLOPT_RETURNTRANSFER => true,
  CURLOPT_POST => true,
  CURLOPT_HTTPHEADER => [
    "Content-Type: application/json",
    "Accept: application/json"
  ],
  CURLOPT_POSTFIELDS => $payload,
  CURLOPT_TIMEOUT => 30,
  CURLOPT_CONNECTTIMEOUT => 10,
  CURLOPT_FOLLOWLOCATION => true,
  CURLOPT_MAXREDIRS => 5,
  CURLOPT_USERAGENT => "FluxFilmHostingerProxy/1.0",
]);

$res = curl_exec($ch);

if ($res === false) {
  $err = curl_error($ch);
  $errno = curl_errno($ch);
  curl_close($ch);
  http_response_code(502);
  echo json_encode(["ok" => false, "message" => "Proxy curl error", "errno" => $errno, "detail" => $err]);
  exit;
}

$code = curl_getinfo($ch, CURLINFO_HTTP_CODE);
curl_close($ch);

// Ensure JSON
$trim = ltrim($res);
if ($trim === "" || ($trim[0] !== "{" && $trim[0] !== "[")) {
  http_response_code(502);
  echo json_encode([
    "ok" => false,
    "message" => "Non-JSON from Apps Script",
    "httpCode" => $code,
    "raw_prefix" => substr($res, 0, 250)
  ]);
  exit;
}

// ── Save to cache if cacheable ──
if (in_array($action, $CACHEABLE) && $res) {
  if (!is_dir($CACHE_DIR)) mkdir($CACHE_DIR, 0755, true);
  file_put_contents($CACHE_DIR . $action . '.json', $res);
}

http_response_code($code ?: 200);
echo $res;
