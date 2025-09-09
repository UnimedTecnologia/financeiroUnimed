<?php

class FormataData {
  
   /**********************************************
        Envia a data no formato DD-MON-RR e
          retorna no formato DD-MM-yyyy
   **********************************************/
   function getData($data){

      $info = array('JAN','FEB','MAR','APR','MAY','JUN','JUL','AUG','SEP','OCT','NOV','DEC');
      $meses = 0;
      while ($meses < 12){
         if ($info[$meses]==substr($data,3,3) ){
           return substr($data,0,2)."/".($meses+1)."/".substr($data,7,2);
         }
         $meses++;
      }
      return "$data";
   }
}

?>
