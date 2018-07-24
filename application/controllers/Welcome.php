<?php
defined( 'BASEPATH' ) || exit( 'No direct script access allowed' );

class Welcome extends CI_Controller
{

    public function index()
    {
        $this->load->helper( 'url' );
        $this->load->helper( 'html' );

        echo anchor( site_url( "welcome/readExcel" ), 'Read Excel', 'title="Read Excel"' );
        echo br( 1 );
        echo anchor( site_url( "welcome/createExcel" ), 'Create Excel', 'title="Create Excel"' );
    }

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

    public function createExcel()
    {
        //load our new PHPExcel library
        $this->load->library( 'excel' );
        //activate worksheet number 1
        $this->excel->setActiveSheetIndex( 0 );
        //name the worksheet
        $this->excel->getActiveSheet()->setTitle( 'test worksheet' );

        //set up the style in an array
        $th_style = array(
            'font'    => array(
                'size'  => 10,
                'bold'  => true,
                'color' => array( 'rgb' => 'ff0000' ),
            ),
            'borders' => array(
                'allborders' => array(
                    'style' => PHPExcel_Style_Border::BORDER_THIN,
                ),
            ),
        );

        $td_style = array(
            'borders' => array(
                'allborders' => array(
                    'style' => PHPExcel_Style_Border::BORDER_THIN,
                ),
            ),
        );

        //apply the style on column A row 1 to Column B row 1
        $this->excel->getActiveSheet()->getStyle( 'A1:D1' )->applyFromArray( $th_style );

        //set cell A1 content with some text
        $this->excel->getActiveSheet()->setCellValue( 'A1', 'Student ID' );
        $this->excel->getActiveSheet()->setCellValue( 'B1', 'Student Name' );
        $this->excel->getActiveSheet()->setCellValue( 'C1', 'Point' );
        $this->excel->getActiveSheet()->setCellValue( 'D1', 'Grade' );

        $student_arr = array(
            array(
                "id"    => "1100",
                "name"  => "Taveevut Nakomah",
                "point" => "80",
                "grade" => "B+",
            ),
            array(
                "id"    => "1101",
                "name"  => "Sakarin Nakomah",
                "point" => "90",
                "grade" => "A",
            ),
            array(
                "id"    => "1102",
                "name"  => "Maple Lovely",
                "point" => "95",
                "grade" => "A",
            ),
            array(
                "id"    => "1103",
                "name"  => "Tuck Boriboon",
                "point" => "55",
                "grade" => "B",
            ),
        );

        $r = 2;
        foreach ( $student_arr as $value ) {
            $this->excel->getActiveSheet()->setCellValue( "A$r", $value["id"] )->getStyle( "A$r" )->applyFromArray( $td_style );
            $this->excel->getActiveSheet()->setCellValue( "B$r", $value["name"] )->getStyle( "B$r" )->applyFromArray( $td_style );
            $this->excel->getActiveSheet()->setCellValue( "C$r", $value["point"] )->getStyle( "C$r" )->applyFromArray( $td_style );
            $this->excel->getActiveSheet()->setCellValue( "D$r", $value["grade"] )->getStyle( "D$r" )->applyFromArray( $td_style );
            $r += 1;
        }

        $filename = 'just_some_random_name.xls'; //save our workbook as this file name

        header( 'Content-Type: application/vnd.ms-excel' ); //mime type
        header( 'Content-Disposition: attachment;filename="' . $filename . '"' ); //tell browser what's the file name
        header( 'Cache-Control: max-age=0' ); //no cache

        //save it to Excel5 format (excel 2003 .XLS file), change this to 'Excel2007' (and adjust the filename extension, also the header mime type)
        //if you want to save it as .XLSX Excel 2007 format
        $objWriter = PHPExcel_IOFactory::createWriter( $this->excel, 'Excel5' );
        //force user to download the Excel file without writing it to server's HD
        $objWriter->save( 'php://output' );
    }
}
