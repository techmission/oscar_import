<?php

/* Constants */

// Values for the import
define('DATA_SOURCE', 'Oscar.org.uk');
define('STATUS_PUBLISHED', 1);
define('STATUS_UNPUBLISHED', 0);

define('ROWS_OSCAR', 131); // How many rows of the spreadsheet have data
define('FILENAME_OSCAR', 'oscar_data.xls'); // The name of the file to import

define('MODE_DEBUG', TRUE); // Whether the script is in debug mode

/* Includes */

// Database connection
include("../db_connection/db.inc");

// Third-party Excel processing library
include("../phpExcelReader/Excel/reader.php");

// Execute the import
oscar_import();

/* Main function for import */
function oscar_import() {
  // Connect to database
  $db = connect_to_db('socgraph');

  // Process file
  $records = process_oscar_records();

  /* if(MODE_DEBUG == TRUE) {
    print_r($records);
  } */

  // Insert records
  $result = insert_oscar_records($records);
  
  // Print debugging results if script is in debug mode
  if(MODE_DEBUG == TRUE) {
   print_r($result);
  }
}

/**
 * Turns the oscar records from the spreadsheet
 * into an associative array for import
 */
function process_oscar_records() {
  // Initialize the spreadsheet reader
  $sheet = new Spreadsheet_Excel_Reader();

  $sheet->setOutputEncoding("UTF-8");

  $sheet->read(FILENAME_OSCAR);

  // Read out the records from the spreadsheet into an array
  $records = array();

  for($i = 1; $i <= ROWS_OSCAR; $i++) {
    for($j = 1; $j <= $sheet->sheets[0]['numCols']; $j++) {
      $records[$i][$j] = $sheet->sheets[0]['cells'][$i][$j];
    }
  }
  
  // Rekey the spreadsheet records
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

/* Inserts the oscar records into the tbl_opportunities table */
function insert_oscar_records(array $records) {
  // Set defaults for the result array
  $result = array('success' => FALSE, 'num_attempted' => 0, 'num_imported' => 0);
  
  // If no records passed in, then return
  if(count($records) < 1) {
    return $result;
  }
  
  // Get the starting guid number
  $start_guid = get_start_guid();
  
  // Set the status for records
  if(MODE_DEBUG == TRUE) {
    $status = STATUS_UNPUBLISHED;
  }
  else {
    $status = STATUS_PUBLISHED;
  }
  
  $count = 0;
  $num_imported = 0;
  foreach($records as $record) {
    // Set the guid
    $guid = 'oscar:' . $start_guid;
    
    // Get the org reference id
    $org_id = get_org_reference_id($record['org_name']);
    
    // Set the raw query string
    $qs = <<<EOT
    INSERT INTO tbl_opportunities
      (title, referralurl, org_name, country_name, description, short_description,
        created, source, org_reference_id, guid, source_guid, status)
      VALUES ('%s','%s', '%s', '%s', '%s', '%s', %d, '%s', %d, '%s', '%s', %d)
EOT;
    // Prepare the query with the values
    $qry = sprintf($qs, 
      mysql_real_escape_string($record['title']), 
      mysql_real_escape_string($record['referralurl']), 
      mysql_real_escape_string($record['org_name']),
      mysql_real_escape_string($record['country_name']),
      mysql_real_escape_string($record['description']),
      mysql_real_escape_string($record['description']),
      time(),
      DATA_SOURCE,
      $org_id,
      $guid,
      $guid,
      $status);
    // Print the query if in debug mode  
    if (MODE_DEBUG == TRUE) {
      echo $qry . PHP_EOL;
    }
    // Execute the query and determine whether the insert was successful
    $result = mysql_query($qry);
    if($result) {
      $count = mysql_affected_rows();
    }
    // Increment the guid number for the insert
    $start_guid++;
    // Increment the number of records imported
    $num_imported = $num_imported + $count;
  }

  // Set the values for success if there were records imported
  if($num_imported > 0) {
    $result = array('success' => TRUE, 'num_attempted' => count($records), 'num_imported' => $num_imported);
  }
  return $result;
}

/* Utility function: gets the org reference id from the org name */
function get_org_reference_id($org_name) {
  $org_id = 0;
  // Look up the title in tbl_organizations
  if(!empty($org_name)) {
    $org_name = mysql_real_escape_string($org_name);
    $qry = 'SELECT sid FROM tbl_organizations WHERE title LIKE "%' . $org_name . '%"';
    $result = mysql_query($qry);
    if($result) {
      $sid = mysql_result($result, 0);
    }
    // If the result is numeric then it can be used as a reference id
    if(is_numeric($sid) && $sid > 0) {
      $org_id = $sid;
    }
  }
  return $org_id;
}

/* Utility function: gets the starting guid from the max in the tbl_opportunities table */
function get_start_guid() {
  $start_guid = 1;
  // Get the numeric part of the GUID
  // @see http://stackoverflow.com/questions/5960620
  $qry = 'SELECT CONVERT(SUBSTRING_INDEX(guid, ":", -1), UNSIGNED INTEGER) as guid_num ';
  $qry .= 'FROM tbl_opportunities WHERE guid LIKE "oscar%" ORDER BY guid_num DESC LIMIT 1';
  $result = mysql_query($qry);
  // If the query is valid
  if($result) {
    $max_guid = mysql_result($result, 0); // Get the field from the first row
  }
  // If the maximum guid is greater than the default start guid, increment
  if($max_guid >= $start_guid) {
    $start_guid = $max_guid +1;
  }
  return $start_guid;
}

