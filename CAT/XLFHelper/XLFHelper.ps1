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

function Copy-SourceToTarget {
    param (
        [string]$XLIFFs
    )
<#
<?xml version="1.0" encoding="UTF-8"?>
<xliff version="1.2" xmlns="urn:oasis:names:tc:xliff:document:1.2" xmlns:okp="okapi-framework:xliff-extensions" xmlns:its="http://www.w3.org/2005/11/its" xmlns:itsxlf="http://www.w3.org/ns/its-xliff/" its:version="2.0">
<file original="xl/sharedStrings.xml" source-language="ko-KR" target-language="vi-VN" datatype="x-undefined">
<body>
<group id="P76C545-sg1" resname="Tabelle1">
<group id="P388F82AB-sg1" resname="1">
</group>
<group id="P388F82AB-sg2" resname="2">
<trans-unit id="P147242AB-tu1" resname="Tabelle1!D2" xml:space="preserve">
<source xml:lang="ko-KR">이 과정에서는 CMMS의 작업 주문에 따라 부품을 받는 프로세스를 설정합니다. </source>
<target xml:lang="vi-VN"></target>
</trans-unit>
</group>
</group>
</body>
</file>
<file original="docProps/core.xml" source-language="ko-KR" target-language="vi-VN" datatype="x-undefined">
<body>
</body>
</file>
</xliff>
#>
    foreach ($file in $XLIFFs) {
        # Đọc nội dung của tệp
        [xml]$xmlContent = Get-Content -Path $file.FullName
        
        # Đặt namespace manager để xử lý namespace trong XML
        $namespaceManager = New-Object System.Xml.XmlNamespaceManager($xmlContent.NameTable)
        $namespaceManager.AddNamespace("ns", "urn:oasis:names:tc:xliff:document:1.2")
    
        # Duyệt qua từng phần tử 'trans-unit' trong tệp .xlf
        foreach ($transUnit in $xmlContent.SelectNodes("//ns:trans-unit", $namespaceManager)) {
            $sourceElement = $transUnit.SelectSingleNode("ns:source", $namespaceManager)
            $targetElement = $transUnit.SelectSingleNode("ns:target", $namespaceManager)
            
            if ($null -ne $sourceElement -and $null -ne $targetElement) {
                # Sao chép nội dung từ <source> sang <target>
                $targetElement.InnerText = $sourceElement.InnerText
            }
        }
        
        # Lưu lại tệp với các thay đổi
        $xmlContent.Save($file.FullName)
    }
    
}

function Wait-ForExit {
    [CmdletBinding()]
    param ()

    Write-Host "Press 'Q' key to exit..."
    try {
        do {
            $key = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
        } until ($key.Character -eq 'q' -or $key.Character -eq 'Q')
    }
    finally {
        Write-Host "Exited."
    }
}