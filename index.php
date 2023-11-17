<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <title>Upload and Process Excel</title>
    <link rel="stylesheet" href="styles.css">
</head>
<body>
    <div class="container">
        <h1>Upload Excel File</h1>

        <form action="index.php" method="post" enctype="multipart/form-data">
            <input type="file" name="excel_file" accept=".xls, .xlsx">
            <input type="submit" name="submit" value="Upload">
        </form>
    </div>

    <?php
    if (isset($_POST['submit'])) {
        $uploadedFile = $_FILES['excel_file']['tmp_name'];
        $targetDirectory = 'uploads' . DIRECTORY_SEPARATOR;
        $targetFile = $targetDirectory . $_FILES['excel_file']['name'];
        move_uploaded_file($uploadedFile, $targetFile);
        //echo "<p>Uploaded File: $targetFile</p>";

        $escapedFilePath = escapeshellarg($targetFile);
        $output = shell_exec("C:\\Users\\Administrator\\AppData\\Local\\Programs\\Python\\Python311\\python.exe process_excel.py $escapedFilePath 2>&1");

        $decodedOutput = json_decode($output, true);
        if ($decodedOutput !== null) {
            echo "<div class='result-container'>";
            echo "<h2>Result:</h2>";
            echo "<table border='1'>";
            foreach ($decodedOutput as $key => $value) {
                $formattedValue = number_format($value);
                // Add the 'number' class to align numbers to the right
                echo "<tr><td>$key</td><td class='number'>$formattedValue</td></tr>";
            }
            echo "</table>";
            echo "</div>";

            unlink($targetFile);

        } else {
            echo "<p>Error decoding JSON output from Python</p>";
        }
    }
    ?>
</body>
</html>
