<?php
session_start();
$token = $_GET['token'] ?? '';
$tokens = $_SESSION['tokens'] ?? [];

if (!isset($tokens[$token])) { http_response_code(404); die('Файл не найден'); }
$f = $tokens[$token];

if (time() > $f['exp'] || !file_exists($f['path'])) {
    unset($tokens[$token]); $_SESSION['tokens'] = $tokens;
    http_response_code(410); die('Ссылка истекла или файл удален');
}

header('Content-Type: application/octet-stream');
header('Content-Disposition: attachment; filename="' . basename($f['name']) . '"');
header('Content-Length: ' . $f['size']);
readfile($f['path']);

// Удаляем файл и токен после скачивания
unlink($f['path']);
unset($tokens[$token]);
$_SESSION['tokens'] = $tokens;
exit;