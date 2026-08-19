<?php
    $conn = new mysqli('localhost', 'root', '', 'work_shop');

    if ($conn->connect_error) {
        die("Connection failed: " . $conn->connect_error);
    }

    $conn->set_charset('utf8');

    date_default_timezone_set('Asia/Bangkok');

    echo "<script>alert('เชื่อมต่อฐานข้อมูลสำเร็จ!');</script>";
?>