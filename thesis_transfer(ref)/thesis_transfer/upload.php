<?php
if ($_FILES["tfile"]["error"] > 0) {
	echo "Error: " . $_FILES["tfile"]["error"];
}
else {
	$command = escapeshellcmd('python3 trans2meta.py ' . $_FILES["tfile"]["tmp_name"]);
	$output = shell_exec($command);
	echo $output;
	$file = "transfered.xlsx";
	header('Content-Type: application/vnd.ms-excel; charset=UTF-8');
	header("Content-Disposition: attachment; filename=" . $file); 
	ob_clean();
	flush();
	readfile($file);
	echo "\nsuccess\n";
	$clean = escapeshellcmd('rm transfered.xlsx');
	$ot = shell_exec($clean);
	echo $ot;
}
?>
<br><br><a href="http://140.118.33.17">返回首頁</a>