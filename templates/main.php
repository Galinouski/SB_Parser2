<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <meta http-equiv="X-UA-Compatible" content="ie=edge">
    <link rel="stylesheet" href="./css/styles.css">
    <title>XLS Parser (D. Galinouski)</title>
</head>
<body>

<div>
    <h2>Парсинг данных excel документа</h2>
    <br>
    <h3>введите данные для парсинга:</h3>
    <br>
    <form method="post" enctype="multipart/form-data" action="./index.php">
        *.XLSX <input type="file" name="file"  />&nbsp;&nbsp;
        <br><br>(Min) начальная цена: <input type="text" name="startPrice" /><br>
        <br>(Max) максимальная цена: <input type="text" name="highPrice" /><br>
        <br>Наименование товара: <input type="text" name="name" /><br>
        <br>В наименовании товара присутствует текст: <input type="text" name="nameLike" /><br>
        <br>Диапазон по артиклю от: <input type="text" name="idStart" /> до: <input type="text" name="idFinish" /><br>

        <br>или начинается с: <input type="text" name="idTitleBegin" /> и оканчивается на: <input type="text" name="idTitleAnd" /><br>
        <br>количество строк парсинга (все по умолчанию): <input type="text" name="limit" /><br>

        <br><br><input type="submit" value="Старт" /><br>
    </form>
</div>
</body>
</html>
