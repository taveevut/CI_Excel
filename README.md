# นำไลบราลี PHPExcel มาใช้งานร่วมกับ CodeIgniter
ดาวน์โหลดไฟล์ PHPExcel จาก gitHub[https://github.com/PHPOffice/PHPExcel]  แล้วแตกไฟล์ zip ดังกล่าว ก๊อปปี้เฉพาะไฟล์ที่อยู่ในแฟ้ม Classes ไปไว้ที่แฟ้ม application/third_party  ดังรูป
<img src="https://github.com/taveevut/CI_Excel/blob/master/screenshot/ex1.PNG">
จากนั้นให้สร้างไฟล์ php เพื่อครอบไลบราลีตัวนี้ไว้โดยบันทึกไว้ในแฟ้ม application/library เพื่อใช้เป็นตัวกลางในการเรียกใช้งาน PHPExcel ในที่นี้ ผมจะสร้างไฟล์ชื่อ Excel.php  แล้วพิมพ์โค๊ดดังต่อไปนี้ลงไป
```sh
<?php if ( ! defined('BASEPATH')) exit('No direct script access allowed');
  require_once APPPATH."/third_party/PHPExcel.php";
  class Excel extends PHPExcel{
    public function __construct(){
        parent::__construct();
    }
  }
```
เท่านี้เราก็สามารถใช้งาน PHPExcel ร่วมกับ CodeIgniter ได้แล้วครับ สำหรับวิธีใช้งานนั้นทำได้ง่าย ๆดังนี้ครับ
#ตัวอย่างการอ่านไฟล์ Excel
สร้างไฟล์ excel ที่จะใช้อ่าน ในที่นี้ผมสร้างไฟล์ excel ที่เก็บข้อมูลคะแนนและเกรดนักศึกษาดังต่อไนี้
<img src="https://github.com/taveevut/CI_Excel/blob/master/screenshot/ex2.png">
จากนั้นบันทึกไว้ที่ใดที่หนึ่งภายใต้ site ในที่นี้ผมเก็บไว้ที่ /upload ตั้งชื่อว่า student_grade.xls
ขั้นตอนต่อไปสร้างไฟล์ controller เพื่อมาทดสอบ ในที่นี้ผมสร้าง controller ชื่อ test.php วางไว้ที่ /controller โดยมีรายละเอียดของโค๊ดดังนี้
```sh
public function readExcel()
{
    $this->load->library( 'excel' );
    $objReader = PHPExcel_IOFactory::load( 'uploads/student_grade.xls' );
    $sheetData = $objReader->getActiveSheet()->toArray( true, true, true, true );
    echo "<table border=1>";
    foreach ( $sheetData as $data ) {
        echo "<tr>";
        $this->writeColumn( $data['A'] );
        $this->writeColumn( $data['B'] );
        $this->writeColumn( $data['C'] );
        $this->writeColumn( $data['D'] );
        echo "</tr>";
    }
    echo "</table>";
}

public function writeColumn( $data )
{
    echo "<td>$data</td>";
}
```
# แหล่งอ้างอิง
* [https://soowoi.wordpress.com/2014/05/11/using-phpexcel-with-codeigniter/]
* [https://github.com/PHPOffice/PHPExcel]