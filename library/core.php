<?php
//функция подключения шаблона с содержимым $context
function render(string $template, array $context) {
    global $base_path;
    extract($context);
    require $base_path . '.\templates\\' . $template . '.php';
}