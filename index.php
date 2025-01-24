<?php
use Shuchkin\SimpleXLSX;

ini_set('error_reporting', E_ALL);
ini_set('display_errors', true);

//require_once __DIR__.'/simplexlsx/src/SimpleXLSX.php';
require_once 'SimpleXLSX.php';

function excelToIcal($inputFile, $outputFile)
{
    if (!$xlsx = SimpleXLSX::parse($inputFile)) {
        die("❌ Error reading Excel file: " . SimpleXLSX::parseError());
    }

    $data = $xlsx->rows();

    // Check if data exists
    if (empty($data)) {
        die("❌ Error: Excel file is empty or not readable!");
    }

    // Create the iCal file header
      $ical = "BEGIN:VCALENDAR\n";
      $ical .= "PRODID:-//hjelmua//NONSGML v1.0//EN\n";
      $ical .= "VERSION:2.0\n";
      $ical .= "X-WR-CALNAME:IBGO\n";  // Change "Bokningar" to any calendar name you prefer
      $ical .= "X-WR-RELCALID:IBGOBokningar\n";  
      $ical .= "BEGIN:VTIMEZONE\n";
      $ical .= "TZID:Europe/Stockholm\n";
      $ical .= "X-LIC-LOCATION:Europe/Stockholm\n";
      $ical .= "BEGIN:DAYLIGHT\n";
      $ical .= "DTSTART:19960331T020000\n";
      $ical .= "RRULE:FREQ=YEARLY;BYDAY=-1SU;BYMONTH=3\n";
      $ical .= "TZNAME:CEST\n";
      $ical .= "TZOFFSETFROM:+0100\n";
      $ical .= "TZOFFSETTO:+0200\n";
      $ical .= "END:DAYLIGHT\n";
      $ical .= "BEGIN:STANDARD\n";
      $ical .= "DTSTART:19961027T030000\n";
      $ical .= "RRULE:FREQ=YEARLY;BYDAY=-1SU;BYMONTH=10\n";
      $ical .= "TZNAME:CET\n";
      $ical .= "TZOFFSETFROM:+0200\n";
      $ical .= "TZOFFSETTO:+0100\n";
      $ical .= "END:STANDARD\n";
      $ical .= "END:VTIMEZONE\n";

    foreach ($data as $index => $row) {
        if ($index == 0) continue; // Skip header row

        // Ensure required columns exist
        if (!isset($row[3], $row[4], $row[8], $row[9], $row[10], $row[11], $row[0], $row[17], $row[12], $row[5], $row[30], $row[2])) {
            continue;
        }

        $summary = trim($row[3]); // Column D for Summary
        $locationDetail = trim($row[4]); // Column E for additional location details

        // If Column E contains "7-spel", modify the Summary
	// Handle "inomhushall 7-spel" first
        if (stripos($locationDetail, 'inomhushall 7-spel') !== false) {
           $summary .= " inomhushall";
        }
        // Handle "7-spel" if "inomhushall 7-spel" was not found
        elseif (stripos($locationDetail, '7-spel') !== false) {
           $summary .= " halvplan";
        }

        // Handle "5-spel"
        if (stripos($locationDetail, '5-spel') !== false) {
            $summary .= " 5-spelsplan";
        }

        // Set LOCATION (Column D and E separated by ", ")
        $location = trim($row[3]) . "\\, " . trim($row[4]); // Column D (3) + Column E (4)
        // Get required columns for DESCRIPTION
        $bokningsnr = trim($row[0]);  // Column A
        $bokningstyp = trim($row[17]); // Column R
        $nyttjare = trim($row[12]);   // Column M
        $frekvens = trim($row[5]);    // Column F
       	$bokad_av = !empty($row[2]) ? trim($row[2]) : "Ingen speciell bokare"; // Default if missing
	$meddelande = !empty($row[30]) ? trim($row[30]) : "Inget bokningsmeddelande"; // Default if missing AE
	// Escape special characters for .ics
        $meddelande = str_replace(
            ['\\', "\n", "\r", ',', ';'], // Characters to escape
            ['\\\\', '\\n', '', '\,', '\;'], // Escaped versions
        $meddelande
        );

         // Truncate to 75 characters (but avoid cutting off in the middle of a word)
         if (mb_strlen($meddelande) > 75) {
         $meddelande = mb_substr($meddelande, 0, 72) . "..."; // Cut at 72 chars + "..."
         }
        $pris = trim($row[14]); // Column O (adjust if necessary)
        $pris = str_replace(',', '.', $pris); // Convert comma to dot for numeric consistency
        $pris = number_format((float)$pris, 2, '.', ''); // Ensure proper decimal format

        // Construct DESCRIPTION field
        $description = "Bokningsnr: $bokningsnr\\n";
        $description .= "Bokningstyp: $bokningstyp\\n";
        $description .= "Nyttjare: $nyttjare\\n";
        $description .= "Frekvens: $frekvens\\n";
	$description .= "Bokningspris: $pris\\n";
	$description .= "Meddelande: $meddelande\\n";
        $description .= "Bokad av: $bokad_av";

        // Convert to proper datetime format
        $startDate = trim($row[7]); // Column H
        $endDate = trim($row[8]);   // Column I
        $startTime = trim($row[9]); // Column J
        $endTime = trim($row[10]);   // Column K


if (!empty($startDate) && !empty($startTime) && !empty($endDate) && !empty($endTime)) {
    $startTimestamp = strtotime("$startDate $startTime");
    $endTimestamp = strtotime("$endDate $endTime");

    // Check if strtotime() failed
    if ($startTimestamp === false) {
        die("<p>❌ Error: strtotime() failed for Start Date/Time: '$startDate $startTime'</p>");
    }
    if ($endTimestamp === false) {
        die("<p>❌ Error: strtotime() failed for End Date/Time: '$endDate $endTime'</p>");
    }

    if ($endTimestamp < $startTimestamp) {
        $endTimestamp = $startTimestamp + 3600; // Default to 1 hour after start
    }

    $dtStart = date("Ymd\THis", $startTimestamp);
    $dtEnd = date("Ymd\THis", $endTimestamp);


            // Create the event
            $event = "BEGIN:VEVENT\n";
            $event .= "DESCRIPTION:$description\n";
	    $event .= "DTEND:$dtEnd\n";
            $event .= "DTSTART:$dtStart\n";
	    $event .= "LOCATION:" . $location . "\n";  // No `addslashes()`, only manually escaped comma
	    $event .= "SUMMARY:$summary\n";
            $event .= "END:VEVENT\n";

            // Add event to iCal content
            $ical .= $event;
        }
    }

    // Close the iCal file
    $ical .= "END:VCALENDAR";

    // Save the iCal file
    file_put_contents($outputFile, $ical);
    return $outputFile;
}

$downloadLink = "";
if ($_SERVER['REQUEST_METHOD'] === 'POST' && isset($_FILES['excelFile'])) {
    $uploadDir = "uploads/";
    if (!is_dir($uploadDir)) {
        mkdir($uploadDir, 0777, true);
    }

    $uploadedFile = $uploadDir . basename($_FILES['excelFile']['name']);
    move_uploaded_file($_FILES['excelFile']['tmp_name'], $uploadedFile);

    $outputFile = "downloads/hjelmuaibgo.ics";
    if (!is_dir("downloads/")) {
        mkdir("downloads/", 0777, true);
    }

    // Unique .ics file name
    $outputFile = "downloads/hjelmuaibgo_" . time() . ".ics";
	
    $icalFile = excelToIcal($uploadedFile, $outputFile);
    $downloadLink = "<a href='$icalFile' class='btn btn-success'><i class='fas fa-download'></i> Download .ics</a>";
}
?>

<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>IBGO-Excel to iCal Converter by hjelmua</title>
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/bootstrap/5.3.0/css/bootstrap.min.css">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.2/css/all.min.css">
    <style>
        body {
            background: #f8f9fa;
            display: flex;
            justify-content: center;
            align-items: center;
            height: 100vh;
            flex-direction: column;
        }
        .container {
            background: white;
            padding: 30px;
            border-radius: 8px;
            box-shadow: 0px 0px 10px rgba(0, 0, 0, 0.1);
            text-align: center;
        }
        h2 {
            color: #343a40;
        }
        .file-upload {
            border: 2px dashed #6c757d;
            padding: 20px;
            background: #e9ecef;
            border-radius: 8px;
            cursor: pointer;
        }
        .file-upload:hover {
            background: #dee2e6;
        }
    </style>
</head>
<body>

<div class="container">
    <h2><i class="fas fa-calendar-alt"></i> IBGO Excel to iCal Converter by hjelmua</h2>
    <p>Upload your Excel file and get a downloadable .ics calendar file.</p>

    <form action="" method="post" enctype="multipart/form-data" class="mt-3">
        <label class="file-upload">
            <input type="file" name="excelFile" accept=".xlsx" required hidden>
            <i class="fas fa-upload"></i> Click here to select an Excel file
        </label>
        <br><br>
        <button type="submit" class="btn btn-primary"><i class="fas fa-sync-alt"></i> Convert to .ics</button>
    </form>

    <div class="mt-3">
        <?php echo $downloadLink; ?>
    </div>
<div class="fixed-bottom">
	<a href="https://github.com/hjelmua/IBGO-Excel_to_iCal" class="btn btn-light btn-sm"><i class="fa-brands fa-github"></i> open source code by hjelmua</a>
</div>
</div>

</body>
</html>
