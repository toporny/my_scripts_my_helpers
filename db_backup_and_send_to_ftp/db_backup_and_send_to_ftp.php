<?php

// 1. script creates SQL dump from databse
// 2. pack this file 
// 3. and send to FTP

define ('FTP_SERVER', 'xxx');
define ('FTP_LOGIN', 'yyy');
define ('FTP_PASSWORD', 'zzz');
define ("LOCAL_BACKUP_PATH", "c:\\BACKUP\\SQL\\");


$database_to_backup = array(
	array("host" => "uuu",  "user" => "vvv",  "pass" =>"xxx",  "db"=>"yyy",  "ftp_dest" => "/db"),
	array("host" => "uuu",  "user" => "vvv",  "pass" =>"xxx",  "db"=>"yyy",  "ftp_dest" => "/db"),
	array("host" => "uuu",  "user" => "vvv",  "pass" =>"xxx",  "db"=>"yyy",  "ftp_dest" => "/db"),
	array("host" => "uuu",  "user" => "vvv",  "pass" =>"xxx",  "db"=>"yyy",  "ftp_dest" => "/db"),
	array("host" => "uuu",  "user" => "vvv",  "pass" =>"xxx",  "db"=>"yyy",  "ftp_dest" => "/db"),
	array("host" => "uuu",  "user" => "vvv",  "pass" =>"xxx",  "db"=>"yyy",  "ftp_dest" => "/db")
);


foreach ($database_to_backup as $item) 
{
	$file_name = date('Y-m-d_His')."_".$item['db'].".sql";
	$command = "mysqldump -h ".$item['host']." -u ".$item['user']." -p".$item['pass']." ".$item['db']." > ".LOCAL_BACKUP_PATH.$file_name;
	// print $command."\n";
	$result = exec ( $command, $output, $return_var);

	if ($return_var !=0) {
		write_log('Error with: '. $command);
	}
	if ($return_var == 0) {
		pack_to_zip($file_name);  // packing....
		if (is_archive_ok($file_name)) {
			unlink (LOCAL_BACKUP_PATH.$file_name); // remove SQL dump
			if (push_to_ftp($file_name.".zip", $item['ftp_dest'])) { // sending ZIP to FTP server
				// unlink (LOCAL_BACKUP_PATH.$file_name.".zip");  // remove locl SQL.ZIP
			}
		}
	}
}

exit;


function pack_to_zip($file_name) {
	$zip = new ZipArchive();
	if ($zip->open(LOCAL_BACKUP_PATH.$file_name.".zip", ZipArchive::CREATE)!==TRUE) {
		$msg = "FATAL: blad podczas tworzenia pliku ZIP ".$file_name."\n";
		write_log($msg);
	} else {
		$zip->addFile(LOCAL_BACKUP_PATH.$file_name);
	}
	$zip->close();
}



function is_archive_ok($file_name) {
	$sql_file_length = filesize(LOCAL_BACKUP_PATH.$file_name);
	$za = new ZipArchive(); 
	$za->open(LOCAL_BACKUP_PATH.$file_name.".zip"); 
	for( $i = 0; $i < $za->numFiles; $i++ ){ 
		$stat = $za->statIndex($i); 
		if (basename( $stat['name'] ) == $file_name) {
			$sql_file_length_inside_zip = $stat['size'];
			if ($sql_file_length == $sql_file_length_inside_zip) {
				return true;
			}
		}
		$msg = "FATAL: Dlugosc pliku sql wewnatrz (".LOCAL_BACKUP_PATH.$file_name.".sql.zip) nie zgadza sie z dlugoscia pliku ".LOCAL_BACKUP_PATH.$file_name;
		write_log($msg);
		return false;
	}
}


function push_to_ftp($file_name, $ftp_destination_dir) {
	$file = LOCAL_BACKUP_PATH.$file_name;
	$remote_file = $ftp_destination_dir."/".$file_name;

	$conn_id = ftp_connect(FTP_SERVER);
	$login_result = ftp_login($conn_id, FTP_LOGIN, FTP_PASSWORD);
	if (ftp_put($conn_id, $remote_file, $file, FTP_BINARY)) {

		$ftp_zip_size = ftp_size ( $conn_id , $remote_file );
		$zip_local_size = filesize($file);

		if ($ftp_zip_size == $zip_local_size) {
			$return = true;	
		} else {
			write_log("FATAL: Dlugosc pliku ZIP na serwerze FTP jest inna niz dlugosc pliku ZIP na lokalnym dysku (".$file.")");
			$return = false;
		}
	} else {
		write_log("FATAL: There was a problem while FTP uploading $file\n");
		$return = false;
	}
	ftp_close($conn_id);
	return $return;
}



function write_log($log_message) {
	file_put_contents(__FILE__.".log", date('Y-m-d_His')." ".$log_message."\n", FILE_APPEND);
}
