<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <meta http-equiv="X-UA-Compatible" content="ie=edge">
    <link rel="stylesheet" href="./css/styles.css">
    <title>XLS Parser (D. Galinouski) Results</title>
</head>
<body>

<h1>Результаты поиска:</h1>

<?php echo $htmlShow; ?>

<br><b>Загрузка в базу данных 'original': <?= $elapsedTime[0] ?> с.</b><br>
<br><b>Загрузка в базу данных 'research': <?= $elapsedTime[1] ?> с.</b><br>

<br>
<br>
<a href="./index.php" class="back">вернуться назад</a>
<br>
</body>
</html>
