<?php
$base_path = __DIR__ . '\\';
//require_once $base_path . 'configs\config.php';
require_once $base_path . 'library\functions.php';

// Подключение шаблона
$context = [];
render('main', $context);