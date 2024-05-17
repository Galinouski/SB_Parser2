<?php

if (!isset($DBConfig)) {
    exit();
}

// Соединение с базой MySQL

$host = $DBConfig['DB_HOST'];
$db = $DBConfig['DB_NAME'];
$charset = $DBConfig['DB_CHARSET'];
$user = $DBConfig['DB_USER'];
$pass = $DBConfig['DB_PASSWORD'];

$dsn = "mysql:host=$host;dbname=$db;charset=$charset";
$opts= [
    \PDO::ATTR_ERRMODE            => \PDO::ERRMODE_EXCEPTION,
    \PDO::ATTR_DEFAULT_FETCH_MODE => \PDO::FETCH_ASSOC,
    \PDO::ATTR_EMULATE_PREPARES   => false,
];

// подключение к базе

try {
    $pdo = new PDO($dsn, $user, $pass, $opts);
} catch (PDOException $e) {
    die($e->getMessage());
}
