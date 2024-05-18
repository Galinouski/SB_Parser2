<?php
global $base_path;
global $pdo;

require_once $base_path ."vendor\autoload.php";
require_once $base_path ."configs\database_config.php";
require_once $base_path ."library\database.php";
require_once $base_path . 'library\functions.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Style\Border;
use PhpOffice\PhpSpreadsheet\Style\Font;


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
            render('errors', $context);
            exit();
        }
    }

    $spreadsheet = new Spreadsheet();
    $spreadsheet = readingXls($fileName);
    $spreadsheet->setActiveSheetIndex(0);
    $sheet = $spreadsheet->getActiveSheet();

    // замеряем время работы загрузки в базу данных
    $startTime = microtime(true);
    // запускаем экспорт данных
    excel2db($spreadsheet, $pdo, 'original');
    $elapsedTime[] = round(microtime(true) - $startTime, 4);

    $sql = "SELECT * FROM `original` WHERE price>0 ";

    if ($name!='' && $nameLike=='')
        $sql .= " AND name = '$name' ";
    if ($nameLike!='' && $name=='')
        $sql .= " AND name LIKE '%$nameLike%' ";
    if ($nameLike!='' && $name!='')
        $sql .= " AND name = '$name' OR name LIKE '%$nameLike%' ";
    if ($idTitleBegin!='' && $idTitleAnd=='')
        $sql .= " AND id LIKE '%$idTitleBegin%' ";
    if ($idTitleAnd!='' && $idTitleBegin=='')
        $sql .= " AND id LIKE '%$idTitleAnd%' ";
    if ($idTitleBegin=='' && $idTitleAnd!='')
        $sql .= " AND id BETWEEN '$idTitleBegin' AND '$idTitleAnd' ";
    if ($idStart!='' && $idFinish=='')
        $sql .= " AND id > '$idStart' ";
    if ($idFinish!='' && $idStart=='')
        $sql .= " AND id <= '$idFinish' ";
    if ($idFinish!='' && $idStart!='')
        $sql .= " AND id BETWEEN '$idStart' AND '$idFinish' ";
    if ($startPrice!='')
        $sql .= " AND price > $startPrice ";
    if ($highPrice!='')
        $sql .= " AND price <= $highPrice ";
    if ($limit!='')
        $sql .= " LIMIT $limit";

    $stmt = $pdo->prepare($sql);
    $stmt->execute();
    $sqlResultArray = $stmt->fetchAll(PDO::FETCH_ASSOC);

    if(!$sqlResultArray){
        // Подключение шаблона errors
        $context = ['badResearch'=>'1'];
        render('errors', $context);
        exit();
    }

    // старт работы c новым xls файлом в котором будет храниться выборка
    $spreadsheet = new Spreadsheet();

    $activesheet = $spreadsheet->getActiveSheet();
    $activesheet->setTitle('Данные выборки');
    $activesheet->getTabColor()->setRGB('FF0000');
    $activesheet->getStyle('A1:D1')->applyFromArray([
        'font' => [
            'name' => 'Arial',
            'bold' => true,
            'italic' => false,
            'underline' => Font::UNDERLINE_DOUBLE,
            'strikethrough' => false,
            'color' => [
                'rgb' => 'red'
            ]
        ],
        'borders' => [
            'allBorders' => [
                'borderStyle' => Border::BORDER_THIN,
                'color' => [
                    'rgb' => 'black'
                ]
            ],
        ],
        'alignment' => [
            'horizontal' => Alignment::HORIZONTAL_CENTER,
            'vertical' => Alignment::VERTICAL_CENTER,
            'wrapText' => true,
        ]
    ]);

    //Вставка данных выборки в xls файл
    pushActiveSheet($activesheet, $sqlResultArray);

    //сохранить лист с выборкой в файл research.xls
    savingXls($spreadsheet, 'research.xls');

    // чтение файла 'research.xls'
    $spreadsheet = readingXls('research.xls');
    $spreadsheet->setActiveSheetIndex(0);
    $sheet = $spreadsheet->getActiveSheet();

    // сохранить в базу данных результат выборки
    // замеряем время работы скрипта
    $startTime = microtime(true);
    excel2db($spreadsheet, $pdo, 'research');
    $elapsedTime[] = round(microtime(true) - $startTime, 4);

    // вывод данных в шаблон results.php
    $html = showsqlResultArrays($sheet);
    $context = ['htmlShow' => $html, 'elapsedTime'=> $elapsedTime];
    render('results', $context);

    $dsn = NULL;
    $_FILES['file']['tmp_name'] = "";



