<?php
	/**
		Related file for Database configuration
		***************************************
		By DarkCoderSc.
	*/

	//{Global variables accessible even by functions}
	$GLOBALS["dbhost"] = "localhost"; // DATABASE HOST DNS/IP
	$GLOBALS["dbname"] = "casperspy"; // DATABASE NAME
	$GLOBALS["dbuser"] = "root";      // DATABASE USERNAME
	$GLOBALS["dbpass"] = "toor";          // DATABASE PASSWORD
	
	/**
		Start a new connection to database
	*/
	function getDBSock(){
		$sock = mysql_connect($GLOBALS["dbhost"],$GLOBALS["dbuser"],$GLOBALS["dbpass"]); 
		if ( ! $sock ) die ("Connection error"); 
		mysql_select_db($GLOBALS["dbname"], $sock) or die ("Connection error");
		return $sock; // Return our socket handle 
	}
	
?>