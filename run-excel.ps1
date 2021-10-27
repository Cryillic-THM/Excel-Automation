function Execute_Macro {

    $Excel = new-object -comobject excel.application

    $Excel_Files = Get-ChildItem -Path C:\Share\ -Include *.xls, *.xlsm, *.xlsx -Recurse

    Foreach($file in $Excel_Files)
    {
       $file.fullname
       $workbook = $excel.workbooks.open($file.fullname)
       $worksheet = $workbook.worksheets.item(1)
       $excel.Run("Macro")
       $workbook.save()
       $workbook.close()
    }
    $excel.quit()
    Remove-Item $file.fullname
}

Execute_Macro
