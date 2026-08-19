<?php 
   require_once("../php/connect.php");
   session_start();
   if( !isset($_SESSION['authen_id'] ) ){
      header('Location: ../../login.php');  
   }  
?>