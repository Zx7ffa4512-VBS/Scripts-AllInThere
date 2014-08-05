<?php
if($_SERVER["REQUEST_METHOD"]=="POST")
{
	$somecontent = file_get_contents("php://input")."\r\n";
	$filename = substr($somecontent,0,strpos($somecontent,'=')).".txt";
	if (!$handle = fopen($filename, 'a+'))
	{
		echo "Can not open file $filename";
		exit;
	}
	if(fwrite($handle,$somecontent) === FALSE)
	{
		echo "Fail";
	}
	else
	{
		echo "OK";
	}
}
?>