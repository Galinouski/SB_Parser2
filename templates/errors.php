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

<?php
if (isset($badResearch)){
    echo "<br><br><span class='error'>К сожалению, ничего не найдено.</span>";
}
else{
    foreach ($errors as $err) {
        echo "<span class='error'>$err</span><br>";
    }
}

?>

<br>
<br>
<a href="./index.php" class="back">вернуться назад</a>
<br>
</body>
</html>