<?php

if (empty($_POST)){
    // Подключение шаблона
    $context = [];
    render('main', $context);
}else{
    // валидация введённых данных
    $errors = [];
    if(!empty($_POST)){

        if ($_FILES['file']['tmp_name']) {
            if (mime_content_type($_FILES['file']['tmp_name']) == 'application/vnd.ms-excel'){
                $fileName = $_FILES['file']['tmp_name'];
            }else $errors[] = "для корректной работы требуется .xls файл.";
        }else $errors[] = "не выбран файл для парсинга.";

        if ($_POST['startPrice']) {
            if(preg_match('/^[\d,\.]+(\d)?$/', $_POST['startPrice'], $Result) === 1) {
                $startPrice = htmlspecialchars($_POST['startPrice'], ENT_QUOTES);
            }else $errors[] = "проверьте поле начальной цены";
        }
        else $startPrice = '';

        if ($_POST['highPrice']) {
            if(preg_match('/^[\d,\.]+(\d)?$/', $_POST['highPrice'], $Result) === 1) {
                $highPrice = htmlspecialchars($_POST['highPrice'], ENT_QUOTES);
            }else $errors[] = "проверьте поле максимальной цены";
        }
        else $highPrice = '';

        if ($_POST['limit']) {
            if(preg_match('/^\d+(\.)?$/', $_POST['limit'], $Result) === 1) {
                $limit = htmlspecialchars($_POST['limit'], ENT_QUOTES);
            }else $errors[] = "проверьте поле количества строк парсинга";
        }
        else $limit = '';

        if ($_POST['idStart']) {
            $idStart = htmlspecialchars($_POST['idStart'], ENT_QUOTES);
        }
        else $idStart = '';

        if ($_POST['idFinish']) {
            $idFinish = htmlspecialchars($_POST['idFinish'], ENT_QUOTES);
        }
        else $idFinish = '';

        if ($_POST['idTitleBegin']) {
            $idTitleBegin = htmlspecialchars($_POST['idTitleBegin'], ENT_QUOTES);
        }
        else $idTitleBegin = '';

        if ($_POST['idTitleAnd']) {
            $idTitleAnd = htmlspecialchars($_POST['idTitleAnd'], ENT_QUOTES);
        }
        else $idTitleAnd = '';

        if ($_POST['name']) {
            $name = htmlspecialchars($_POST['name'], ENT_QUOTES);
        }
        else $name = '';

        if ($_POST['nameLike']) {
            $nameLike = htmlspecialchars($_POST['nameLike'], ENT_QUOTES);
        }
        else $nameLike = '';

        if(!empty($errors)){

            // Подключение шаблона errors
            $context = ['errors'=>$errors];
            render('main', $context);
            exit();
        }
    }
    require_once $base_path . 'baseController.php';
}