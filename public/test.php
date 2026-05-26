<?php
error_reporting(E_ALL);
ini_set('display_errors', 1);

echo "Тест подключения библиотек:<br>";

// Проверка vendor/autoload.php
$autoloadPath = __DIR__ . '/../vendor/autoload.php';
echo "Путь к autoload: $autoloadPath<br>";
echo "Файл существует: " . (file_exists($autoloadPath) ? 'ДА' : 'НЕТ') . "<br><br>";

if (file_exists($autoloadPath)) {
    try {
        require_once $autoloadPath;
        echo "✅ Autoload подключен<br>";
        
        // Проверка TCPDF
        if (class_exists('TCPDF')) {
            echo "✅ TCPDF доступен<br>";
        } else {
            echo "❌ TCPDF НЕ найден<br>";
        }
        
        // Проверка PhpWord
        if (class_exists('PhpOffice\PhpWord\PhpWord')) {
            echo "✅ PhpWord доступен<br>";
        } else {
            echo "❌ PhpWord НЕ найден<br>";
        }
        
        // Проверка PhpSpreadsheet
        if (class_exists('PhpOffice\PhpSpreadsheet\IOFactory')) {
            echo "✅ PhpSpreadsheet доступен<br>";
        } else {
            echo "❌ PhpSpreadsheet НЕ найден<br>";
        }
        
    } catch (Exception $e) {
        echo "❌ Ошибка при подключении: " . $e->getMessage();
    }
} else {
    echo "❌ Файл vendor/autoload.php не найден!<br>";
    echo "Выполните в терминале: composer install<br>";
}

echo "<br><br>Структура папок:<br>";
$baseDir = __DIR__ . '/..';
if (is_dir($baseDir . '/vendor')) {
    echo "✅ Папка vendor существует<br>";
    $files = scandir($baseDir . '/vendor');
    echo "Содержимое: " . implode(', ', array_filter($files, fn($f) => $f[0] !== '.'));
} else {
    echo "❌ Папка vendor НЕ существует<br>";
}
?>