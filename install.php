<?php
$config = require 'config.php';
$mysqli = new mysqli($config['host'], $config['user'], $config['pass']);
if ($mysqli->connect_error) die('Ошибка подключения к MySQL');

$sql = file_get_contents(__DIR__ . '/database.sql');
$mysqli->multi_query($sql);
echo 'База данных и таблицы созданы/обновлены. Можно удалять install.php';
?>