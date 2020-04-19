# phpword

```php
		
		$PHPWord = new \PhpOffice\PhpWord\PhpWord();
        $section = $PHPWord->addSection();
        $header = $section->createHeader();
        $table = $header->addTable();
        $table->addRow();
//        $imgUrl = 'http://www.baidu.com/logowhie.png';
        $imgUrl = base_url('/media/logo.png');
        $cell = $table->addCell(9000);
        $cell->addImage($imgUrl, ['width' => 200,'height' => 50,'align' => 'left']);
        $PHPWord->addTitleStyle(2, ['bold' => true, 'size' => 20, 'name' => 'Times New Roman', 'Color' => '#1F2D3C'], ['align' => 'center']);
        $section->addTitle("标题", 2);
        $styleTable = ['borderSize'=> 16, 'borderColor'=>'#3B5E82', 'cellMargin'=>80];
        $fontCnStyle = ['bold' => true, 'valign' => 'center','color' => '#000','size' => 12,'name' => '宋体'];
        $fontEnStyle = ['bold' => true, 'valign' => 'center','color' => '#000','size' => 12,'name' => 'Times New Roman'];
        $PHPWord->addTableStyle('table_1',$styleTable);
        $table = $section->addTable('table_1');
        foreach($result as $vv){
            $table->addRow();
            $cell =  $table->addCell(9000);
           if($vv->NameE){
               $cell->addText('Name： '.$vv->NameE,$fontEnStyle);
               $cell->addText(' ' );
           }

           if($vv->PostE){
               $cell->addText('Post： '.$vv->PostE,$fontEnStyle);
               $cell->addText(' ');
           }

           if($vv->UnitE){
               $cell->addText('Company： '.$vv->UnitE,$fontEnStyle);
               $cell->addText(' ');
           }

           if($vv->NameC){
               $cell->addText('姓名： '.$vv->NameC,$fontCnStyle);
               $cell->addText(' ');
           }
           if($vv->Post){
               $cell->addText('公司&职务： '.$vv->Post,$fontCnStyle);
               $cell->addText(' ');
           }

           $cell->addText($vv->Email,$fontEnStyle);
        }

//        $filename = 'xxx'.time();
        $objWrite = \PhpOffice\PhpWord\IOFactory::createWriter($PHPWord, 'Word2007');
//        $objWrite->save(FILEPATH.'word/'.$filename.'.docx');
        $filename ='xxx.docx';
        header("Pragma: public");
        header("Expires: 0");
        header('Access-Control-Allow-Origin:*');
        header('Access-Control-Allow-Headers:content-type');
        header('Access-Control-Allow-Credentials:true');
        header("Cache-Control:must-revalidate, post-check=0, pre-check=0");
        header("Content-Type:application/force-download");
        header("Content-Type:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        header("Content-Type:application/octet-stream");
        header("Content-Type:application/download");;
        header("Content-Disposition:attachment;filename=$filename");
        header("Content-Transfer-Encoding:binary");
        $objWrite->save('php://output');
        exit();
```

