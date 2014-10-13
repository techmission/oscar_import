<?php

/* Constants */

// Values for the import
define('DATA_SOURCE', 'Oscar.org.uk');
define('STATUS_PUBLISHED', 1);
define('STATUS_UNPUBLISHED', 0);

define('ROWS_OSCAR', 131); // How many rows of the spreadsheet have data
define('FILENAME_OSCAR', 'oscar_data.xls'); // The name of the file to import

define('MODE_DEBUG', 1); // Whether the script is in debug mode

/* Includes */

// Database connection
include("../db_connection/db.inc");

// Third-party Excel processing library
include("../phpExcelReader/Excel/reader.php");

// Execute the import
oscar_import();

// Main function for import
function oscar_import() {
  // Connect to database
  $db = connect_to_db('socgraph');

  // Process file
  $records = process_oscar_records();

  if(MODE_DEBUG == TRUE) {
    print_r($records);
  }

  // Insert records
  $result = insert_oscar_records($records);
  
  // Print debugging results if script is in debug mode
  if(MODE_DEBUG == TRUE) {
   print_r($result);
  }
}

// Turn the oscar records into an associative array for import
function process_oscar_records() {
  $sheet = new Spreadsheet_Excel_Reader();

  $sheet->setOutputEncoding("UTF-8");

  $sheet->read(FILENAME_OSCAR);

  $records = array();

  for($i = 1; $i <= ROWS_OSCAR; $i++) {
    for($j = 1; $j <= $sheet->sheets[0]['numCols']; $j++) {
      $records[$i][$j] = $sheet->sheets[0]['cells'][$i][$j];
    }
  }

  $key = '';
  foreach($records as $rownum => $record) {
    foreach($record as $colnum => $value) {
      switch($colnum) {
        case 1:
          $key = 'title';
          break;
        case 2:
          $key = 'referralurl';
          break;
        case 3:
          $key = 'org_name';
          break;
        case 4:
          $key = 'country_name';
          break;
        case 5:
          $key = 'description';
          break;
      }
      $records[$rownum][$key] = $records[$rownum][$colnum];
      unset($records[$rownum][$colnum]);
    }
  }
  return $records;
}

// Insert the oscar records into the tbl_opportunities table
function insert_oscar_records(array $records) {
  return array('success' => FALSE, 'num_imported' => 0);
}

