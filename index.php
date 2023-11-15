<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <title>Upload and Process Excel</title>
</head>
<body>
    <h1>Upload Excel File</h1>

    <form action="index.php" method="post" enctype="multipart/form-data">
        <input type="file" name="excel_file" accept=".xls, .xlsx">
        <input type="submit" name="submit" value="Upload">
    </form>

    <?php
    if (isset($_POST['submit'])) {
        $uploadedFile = $_FILES['excel_file']['tmp_name'];
        $targetDirectory = 'uploads' . DIRECTORY_SEPARATOR;  // Use DIRECTORY_SEPARATOR
        $targetFile = $targetDirectory . $_FILES['excel_file']['name'];
        move_uploaded_file($uploadedFile, $targetFile);
        echo "<p>Uploaded File: $targetFile</p>";

        // Use escapeshellarg to properly escape the file path
        $escapedFilePath = escapeshellarg($targetFile);
        $output = shell_exec("C:\\Users\\Administrator\\AppData\\Local\\Programs\\Python\\Python311\\python.exe process_excel.py $escapedFilePath 2>&1");

        // Decode JSON output and display as a table with commas in numbers
        $decodedOutput = json_decode($output, true);
        if ($decodedOutput !== null) {
            echo "<h2>Result:</h2>";
            echo "<table border='1'>";
            foreach ($decodedOutput as $key => $value) {
                $formattedValue = number_format($value);  // Add commas to numbers
                echo "<tr><td>$key</td><td>$formattedValue</td></tr>";
            }
            echo "</table>";

            // Delete the uploaded Excel file
            unlink($targetFile);
            
        } else {
            echo "<p>Error decoding JSON output from Python</p>";
        }
    }
    ?>
</body>
</html>
