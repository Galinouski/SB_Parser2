<?php

use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xls;

/**
 * @param Spreadsheet $spreadsheet - Excel-книга с данными
 * @param string $fileName  - имя xls файла
 * @param PDO $pdo   - PDO-подключение к базе данных
 * @param string $table  - имя таблицы в базе данных
 * @throws \PhpOffice\PhpSpreadsheet\Exception
 */


function readingXls(string $fileName){
    // Чтение xls файл с начальными данными
    //$spreadsheet = new Spreadsheet();
    $reader = IOFactory::createReader('Xls');
    $spreadsheet = $reader->load($fileName);

    return $spreadsheet;
}

function savingXls(Spreadsheet $spreadsheet, string $fileName){
    try {
        $writer = new Xls($spreadsheet);
        $writer->save($fileName);

    } catch (PhpOffice\PhpSpreadsheet\Writer\Exception $e) {
        echo $e->getMessage();
    }
}

function showResults ($sheet): string
{
    // формирование html-кода с данными
    $html = '<br><table style="width: 70%;">';
    foreach ($sheet->getRowIterator() as $row) {
        $html .= '<tr>';
        $cellIterator = $row->getCellIterator();
        foreach ($cellIterator as $cell) {

            // значение текущей ячейки
            $value = $cell->getCalculatedValue();

            $html .= '<td>'.$value.'</td>';
        }
        $html .= '<tr>';
    }
    $html .= '</table>';

    return $html;
}

function excel2db(Spreadsheet $spreadsheet, PDO $pdo, $table)
{
    // получает названия листов книги в виде массива
    $sheetNames = $spreadsheet->getSheetNames();

    // возвращает количество листов в книге
    $sheetsCount = $spreadsheet->getSheetCount();

    // проходимся по каждому листу
    for ($c = 0; $c < $sheetsCount; $c++)
    {
        // ссылка на лист
        $sheet = $spreadsheet->getSheet($c);
        // последняя строка в листе
        $highestRow = $sheet->getHighestRow('A');

        // SQL-запросы на вставку данных в базу

        $query_string = "DROP TABLE IF EXISTS $table ";
        $stmt = $pdo->prepare($query_string);
        $res = $stmt->execute();

        $query_string = "CREATE TABLE IF NOT EXISTS $table (`id` TEXT NULL , `name` TEXT NULL , `price` TEXT NULL , `store` TEXT NULL ) ENGINE = InnoDB ";
        $stmt = $pdo->prepare($query_string);
        $res = $stmt->execute();

        $sql = "INSERT INTO $table (
                               id, name, price, store
                         )
                         VALUES (:id, :name, :price, :store)";

        // подготовленное SQL-выражение
        $stmt = $pdo->prepare($sql);

        // проходимся по каждой строке в листе
        // счетчик начинается с 2-ой строки, так как первая строка - это заголовок
        for ($i = 2; $i < $highestRow + 1; $i++)
        {
            // получаем значения из ячеек столбцов
            $id = $sheet->getCell('A' . $i)->getValue();
            $name = $sheet->getCell('B' . $i)->getValue();
            $price = $sheet->getCell('C' . $i)->getValue();
            $store = $sheet->getCell('D' . $i)->getValue();


            $stmt->bindParam(':id', $id);
            $stmt->bindParam(':name', $name);
            $stmt->bindParam(':price', $price);
            $stmt->bindParam(':store', $store);

            $stmt->execute();
        }
    }
}