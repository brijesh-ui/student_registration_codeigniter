<?php
if (!defined('BASEPATH'))
    exit('No direct script access allowed');

class Export extends CI_Controller {

    // construct
    public function __construct() {
        parent::__construct();
        // load model
         $this->load->library("excel");
        $this->load->model('Export_model','export');

    }
  public function createXLS() {
        // create file name
  $fileName = 'data-'.time().'.xlsx';       
  $this->load->library("excel");
  $student_info = $this->export->studentList();
  $object = new PHPExcel();
  $object->setActiveSheetIndex(0);

  //set header for excel file
  $object->getActiveSheet()->SetCellValue('A1', 'First Name');
  $object->getActiveSheet()->SetCellValue('B1', 'Last Name');
  $object->getActiveSheet()->SetCellValue('C1', 'Email');
  $object->getActiveSheet()->SetCellValue('D1', 'Gender');
  $object->getActiveSheet()->SetCellValue('E1', 'Education');
  $object->getActiveSheet()->SetCellValue('F1', 'Hobbies');

  // set row for excel file
  $rowCount = 2;
        foreach ($student_info as $list) 
        {
            $object->getActiveSheet()->SetCellValue('A' . $rowCount, $list->first_name);
            $object->getActiveSheet()->SetCellValue('B' . $rowCount, $list->last_name);
            $object->getActiveSheet()->SetCellValue('C' . $rowCount, $list->email);
            $object->getActiveSheet()->SetCellValue('D' . $rowCount, $list->gender);
            $object->getActiveSheet()->SetCellValue('E' . $rowCount, $list->education);
            $object->getActiveSheet()->SetCellValue('E' . $rowCount, $list->hobbies);
            $rowCount++;
        }

        $filename = "student_info". date("Y-m-d-H-i-s").".csv";
        header('Content-Type: application/vnd.ms-excel'); 
        header('Content-Disposition: attachment;filename="'.$filename.'"');
        header('Cache-Control: max-age=0'); 
        $objWriter = PHPExcel_IOFactory::createWriter($object, 'CSV');  
        $objWriter->save('php://output'); 
  
 }


     public function import_student_data()
       {
        $this->load->view('admin/import');

       }


// import data into database in add_student table......
     public function import_data()
       {

        if($this->input->post('import'))
        { 
            $path = 'uploads/';
            $config['upload_path'] = $path;
            $config['allowed_types'] = 'xlsx|xls|csv';
            $config['remove_spaces'] = TRUE;
            $this->upload->initialize($config);
            $this->load->library("excel",$config);
            if (!$this->upload->do_upload('uploadFile')) {
                $error = array('error' => $this->upload->display_errors());
            } else {
                $data = array('upload_data' => $this->upload->data());
            }
            if(empty($error))
            {
              if (!empty($data['upload_data']['file_name'])) 
              {
                $import_xls_file = $data['upload_data']['file_name'];
            } 
            else {
                $import_xls_file = 0;
            }
            $inputFileName = $path . $import_xls_file;
            
            try {
                $inputFileType = PHPExcel_IOFactory::identify($inputFileName);
                $objReader = PHPExcel_IOFactory::createReader($inputFileType);
                $objPHPExcel = $objReader->load($inputFileName);
                $allDataInSheet = $objPHPExcel->getActiveSheet()->toArray(null, true, true, true);
                $flag = true;
                $i=0;
                foreach ($allDataInSheet as $value) {
                  if($flag){
                    $flag =false;
                    continue;
                  }
                  $inserdata[$i]['first_name'] = $value['A'];
                  $inserdata[$i]['last_name'] = $value['B'];
                  $inserdata[$i]['email'] = $value['C'];
                  $inserdata[$i]['gender'] = $value['D'];
                  $inserdata[$i]['education'] = $value['E'];
                  $inserdata[$i]['hobbies'] = $value['F'];
                  $i++;
                }               
                $result = $this->export->import_data($inserdata);   
                if($result){
                  echo "Imported successfully";
                }else{
                  echo "ERROR !";
                }             
 
          } catch (Exception $e) {
               die('Error loading file "' . pathinfo($inputFileName, PATHINFO_BASENAME)
                        . '": ' .$e->getMessage());
            }
          }else{
              echo $error['error'];
            }
            
            
    }
    $this->load->view('admin/import');

   } // end function

 
public function import_student_seet()
    {

        if($this->input->post('import'))
        { 
            $path = 'uploads/';
            $config['upload_path'] = $path;
            $config['allowed_types'] = 'xlsx|xls|csv';
            $config['remove_spaces'] = TRUE;
            $this->upload->initialize($config);
            $this->load->library("excel",$config);
            if (!$this->upload->do_upload('uploadFile')) {
                $error = array('error' => $this->upload->display_errors());
            } else {
                $data = array('upload_data' => $this->upload->data());
            }
            if(empty($error))
            {
              if (!empty($data['upload_data']['file_name'])) 
              {
                $import_xls_file = $data['upload_data']['file_name'];
            } 
            else {
                $import_xls_file = 0;
            }
            $inputFileName = $path . $import_xls_file;
            
            try {
                $inputFileType = PHPExcel_IOFactory::identify($inputFileName);
                $objReader = PHPExcel_IOFactory::createReader($inputFileType);
                $objPHPExcel = $objReader->load($inputFileName);
                $allDataInSheet = $objPHPExcel->getActiveSheet()->toArray(null, true, true, true);
                $flag = true;
                $i=0;
                foreach ($allDataInSheet as $value) {
                  if($flag){
                    $flag =false;
                    continue;
                  }

                // get data from database from contact person table........  
                $contact_person = $this->db->where("contact_email", $value['I'])
                                             ->where("contact_name", $value['J'])
                                             ->from("wp_contact_person")
                                             ->get();
                                     $person = $contact_person->result_array();  

                  if(count($person)>0)
                  {
                    $person_id = $person[0]['contact_id'];

                }
                  else
                  {
                    // insert data into contact prso table
                    $query = $this->db->insert("wp_contact_person", array("contact_phone"=>$value['H'], "contact_email"=>$value['I'], "contact_name"=> $value['J'], "contact_relation"=>$value['K']));
                    $person_id = $this->db->insert_id();
                  } // end loop of contact person
               
                  // get data from wp_branch and select branch_code ....
                  $branch = $this->db->where("branch_code ",$value['B'])
                                     ->from("wp_branch")
                                     ->get();
                      $branch_result = $branch->result_array(); 
                   if(count($branch_result)>0)
                  {
                    $branch_code = $branch_result[0]['branch_code'];
                  }
                 // end loop of branch

                 //get data from school_levels and select  level_id ....
                  $level = $this->db->where("level_name", $value['G'])
                                     ->from("wp_school_levels")
                                     ->get();
                      $levet_result = $level->result_array(); 
                   if(count($levet_result)>0)
                  {
                    $level_name = $levet_result[0]['level_name'];
                  }
                  // end loop of level

                  $inserdata[$i]['registration_no'] = $value['A'];
                  $inserdata[$i]['branch_code'] = $value['B'];
                  $inserdata[$i]['student_branch'] = $branch_code;
                  $inserdata[$i]['student_name'] = $value['C'];
                  $inserdata[$i]['address'] = $value['D'];
                  $inserdata[$i]['school_year'] = $value['E'];
                  $inserdata[$i]['school_id'] = $value['F'];
                  $inserdata[$i]['parent_id'] = $person_id;
                  $inserdata[$i]['school_level'] = $value['G'];
                  $inserdata[$i]['contact_phone'] = $value['H'];
                  $inserdata[$i]['contact_email'] = $value['I'];
                  $inserdata[$i]['contact_name'] = $value['J'];
                  $inserdata[$i]['contact_relation'] = $value['K'];
                  $inserdata[$i]['stream_name'] = $value['L'];
                  $inserdata[$i]['student_mobile'] = $value['M'];
                  $i++;

                   $register = $this->db->insert('wp_student_registration', ['registration_no'=>$value['A'], 'student_branch'=>$branch_code, 'student_name'=>$value['C'], 'address'=>$value['D'],
                'school_year'=> $value['E'], 'school_level'=>$value['G'], 'parent_id'=>$person_id, 
                'student_mobile'=>$value['M']]);
                   
               
                  }//end for each loop.........

                if($register)
                {
                  echo "Data Inserted successfully";
                }else
                {
                  echo "Data Not Inserted";
                }


 
          } catch (Exception $e) {
               die('Error loading file "' . pathinfo($inputFileName, PATHINFO_BASENAME)
                        . '": ' .$e->getMessage());
            }
          }else{
              echo $error['error'];
            }
            
            
    }  //end if loop
    $this->load->view('admin/import');

   } // finish tha function























} // end class
        


?>