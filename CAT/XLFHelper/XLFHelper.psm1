function Convert-ExcelToXliff {
    param (
        [string]$excelFilePath,
        [string]$xliffFilePath,
        [string]$sourceCol = "A",
        [string]$targetCol = "B",
        [string]$sourceLang = "en",
        [string]$targetLang = "vi"
    )

    # Mở Excel và tải tệp
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $workbook = $excel.Workbooks.Open($excelFilePath)
    $worksheet = $workbook.Sheets.Item(1)

    # Tạo tài liệu XML
    $xmlDoc = New-Object System.Xml.XmlDocument
    $xmlDecl = $xmlDoc.CreateXmlDeclaration("1.0", "UTF-8", $null)
    $xmlDoc.AppendChild($xmlDecl) | Out-Null

    # Tạo phần tử gốc xliff
    $xliffElem = $xmlDoc.CreateElement("xliff")
    $xliffElem.SetAttribute("version", "1.2")
    $xmlDoc.AppendChild($xliffElem) | Out-Null

    # Tạo phần tử file
    $fileElem = $xmlDoc.CreateElement("file")
    $fileElem.SetAttribute("source-language", $sourceLang)
    $fileElem.SetAttribute("target-language", $targetLang)
    $fileElem.SetAttribute("datatype", "plaintext")
    $fileElem.SetAttribute("original", [System.IO.Path]::GetFileName($excelFilePath))
    $xliffElem.AppendChild($fileElem) | Out-Null

    # Tạo phần tử body
    $bodyElem = $xmlDoc.CreateElement("body")
    $fileElem.AppendChild($bodyElem) | Out-Null

    # Lặp qua các hàng và tạo các phần tử trans-unit
    $rowIndex = 2
    while ($worksheet.Cells.Item($rowIndex, $sourceCol).Text -ne "") {
        $sourceText = $worksheet.Cells.Item($rowIndex, $sourceCol).Text
        $targetText = $worksheet.Cells.Item($rowIndex, $targetCol).Text

        $transUnitElem = $xmlDoc.CreateElement("trans-unit")
        $transUnitElem.SetAttribute("id", $rowIndex.ToString())
        $bodyElem.AppendChild($transUnitElem) | Out-Null

        $sourceElem = $xmlDoc.CreateElement("source")
        $sourceElem.InnerText = $sourceText
        $transUnitElem.AppendChild($sourceElem) | Out-Null

        $targetElem = $xmlDoc.CreateElement("target")
        $targetElem.InnerText = $targetText
        $transUnitElem.AppendChild($targetElem) | Out-Null

        $rowIndex++
    }

    # Lưu tài liệu XML
    $xmlDoc.Save($xliffFilePath)

    # Đóng Excel
    $workbook.Close($false)
    $excel.Quit()

    Write-Host "Converted $excelFilePath to $xliffFilePath"
}

function Update-ExcelFromXliff {
    param (
        [string]$excelFilePath,
        [string]$xliffFilePath,
        [string]$sourceCol = "A",
        [string]$targetCol = "B"
    )

    # Mở Excel và tải tệp
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $workbook = $excel.Workbooks.Open($excelFilePath)
    $worksheet = $workbook.Sheets.Item(1)

    # Mở tệp XLIFF
    $xmlDoc = New-Object System.Xml.XmlDocument
    $xmlDoc.Load($xliffFilePath)

    # Lấy các phần tử trans-unit
    $transUnits = $xmlDoc.SelectNodes("//trans-unit")

    # Lặp qua các trans-unit và cập nhật Excel
    foreach ($transUnit in $transUnits) {
        $id = $transUnit.GetAttribute("id")
        $targetText = $transUnit.SelectSingleNode("target").InnerText

        $rowIndex = [int]$id
        $worksheet.Cells.Item($rowIndex, $targetCol).Value2 = $targetText
    }

    # Lưu và đóng tệp Excel
    $workbook.Save()
    $workbook.Close($false)
    $excel.Quit()

    Write-Host "Updated $excelFilePath with translations from $xliffFilePath"
}

Export-ModuleMember -Function Convert-ExcelToXliff
Export-ModuleMember -Function Update-ExcelFromXliff