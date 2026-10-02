param(
  [Parameter(Mandatory = $true)]
  [ValidateSet("status", "list-open", "excel-batch", "word-batch", "powerpoint-batch")]
  [string]$Command
)

$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest

function Release-ComObject([object]$Value) {
  if ($null -ne $Value -and [Runtime.InteropServices.Marshal]::IsComObject($Value)) {
    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($Value)
  }
}

function Get-ActiveComObject([string]$ProgId) {
  try { return [Runtime.InteropServices.Marshal]::GetActiveObject($ProgId) } catch { return $null }
}

function Convert-HexColor([string]$Hex) {
  if ([string]::IsNullOrWhiteSpace($Hex)) { return $null }
  $value = $Hex.TrimStart('#')
  if ($value.Length -ne 6) { throw "Color must be a six-digit hex value: $Hex" }
  $r = [Convert]::ToInt32($value.Substring(0, 2), 16)
  $g = [Convert]::ToInt32($value.Substring(2, 2), 16)
  $b = [Convert]::ToInt32($value.Substring(4, 2), 16)
  return $r + (256 * $g) + (65536 * $b)
}

function Get-Property($Object, [string]$Name, $Default = $null) {
  $property = $Object.PSObject.Properties[$Name]
  if ($null -eq $property -or $null -eq $property.Value) { return $Default }
  return $property.Value
}

function Read-JsonInput {
  $json = [Console]::In.ReadToEnd()
  if ([string]::IsNullOrWhiteSpace($json)) { return [pscustomobject]@{} }
  return $json | ConvertFrom-Json
}

function Write-JsonOutput($Value) {
  [Console]::Out.WriteLine(($Value | ConvertTo-Json -Depth 30 -Compress))
}

function Get-OfficeStatus {
  $items = @()
  foreach ($definition in @(
    @{ name = "excel"; progId = "Excel.Application" },
    @{ name = "word"; progId = "Word.Application" },
    @{ name = "powerpoint"; progId = "PowerPoint.Application" }
  )) {
    $app = $null
    try {
      $type = [Type]::GetTypeFromProgID($definition.progId)
      $app = Get-ActiveComObject $definition.progId
      $items += [ordered]@{
        application = $definition.name
        installed = $null -ne $type
        running = $null -ne $app
        version = if ($null -ne $app) { [string]$app.Version } else { $null }
      }
    } finally {
      Release-ComObject $app
    }
  }
  return @{ applications = $items }
}

function Get-OpenOfficeFiles {
  $result = @{ excel = @(); word = @(); powerpoint = @() }
  $app = $null
  try {
    $app = Get-ActiveComObject "Excel.Application"
    if ($null -ne $app) {
      foreach ($book in $app.Workbooks) {
        $result.excel += @{ name = [string]$book.Name; path = [string]$book.FullName; active = $book.Name -eq $app.ActiveWorkbook.Name }
        Release-ComObject $book
      }
    }
  } finally { Release-ComObject $app }

  $app = $null
  try {
    $app = Get-ActiveComObject "Word.Application"
    if ($null -ne $app) {
      foreach ($document in $app.Documents) {
        $result.word += @{ name = [string]$document.Name; path = [string]$document.FullName; active = $document.Name -eq $app.ActiveDocument.Name }
        Release-ComObject $document
      }
    }
  } finally { Release-ComObject $app }

  $app = $null
  try {
    $app = Get-ActiveComObject "PowerPoint.Application"
    if ($null -ne $app) {
      foreach ($presentation in $app.Presentations) {
        $result.powerpoint += @{ name = [string]$presentation.Name; path = [string]$presentation.FullName; active = $presentation.Name -eq $app.ActivePresentation.Name }
        Release-ComObject $presentation
      }
    }
  } finally { Release-ComObject $app }
  return $result
}

function Get-ExcelSession($Target) {
  $attach = [bool](Get-Property $Target "active" $false)
  $app = if ($attach) { Get-ActiveComObject "Excel.Application" } else { $null }
  $ownedApp = $null -eq $app
  if ($ownedApp) { $app = New-Object -ComObject Excel.Application }
  if ($ownedApp) { $app.Visible = [bool](Get-Property $Target "visible" $false) }
  if ($ownedApp) { $app.DisplayAlerts = $false }

  $book = $null
  $openedHere = $false
  $targetPath = [string](Get-Property $Target "path" "")
  if ([bool](Get-Property $Target "create" $false)) {
    $book = $app.Workbooks.Add()
    $openedHere = $true
  } elseif (-not [string]::IsNullOrWhiteSpace($targetPath)) {
    foreach ($candidate in $app.Workbooks) {
      if ([string]::Equals([string]$candidate.FullName, $targetPath, [StringComparison]::OrdinalIgnoreCase)) {
        $book = $candidate
        break
      }
      Release-ComObject $candidate
    }
    if ($null -eq $book) {
      $book = $app.Workbooks.Open($targetPath)
      $openedHere = $true
    }
  } else {
    $book = $app.ActiveWorkbook
  }
  if ($null -eq $book) { throw "No Excel workbook is available. Provide path, create=true, or active=true." }
  return @{ app = $app; document = $book; ownedApp = $ownedApp; openedHere = $openedHere }
}

function Get-ExcelRangeSnapshot($Range) {
  $rows = [int]$Range.Rows.Count
  $columns = [int]$Range.Columns.Count
  $values = @()
  $formulas = @()
  for ($row = 1; $row -le $rows; $row++) {
    $valueRow = @()
    $formulaRow = @()
    for ($column = 1; $column -le $columns; $column++) {
      $cell = $Range.Cells.Item($row, $column)
      $valueRow += $cell.Value2
      $formulaRow += if ($cell.HasFormula) { [string]$cell.Formula } else { $null }
      Release-ComObject $cell
    }
    $values += ,$valueRow
    $formulas += ,$formulaRow
  }
  return @{ address = [string]$Range.Address($false, $false); rows = $rows; columns = $columns; values = $values; formulas = $formulas }
}

function Invoke-ExcelBatch($Request) {
  $session = Get-ExcelSession $Request.target
  $app = $session.app
  $book = $session.document
  $results = @()
  try {
    foreach ($operation in $Request.operations) {
      $op = [string]$operation.op
      switch ($op) {
        "inspect" {
          $sheets = @()
          foreach ($sheet in $book.Worksheets) {
            $used = $sheet.UsedRange
            $sheets += @{ name = [string]$sheet.Name; rows = [int]$used.Rows.Count; columns = [int]$used.Columns.Count; usedRange = [string]$used.Address($false, $false); tableCount = [int]$sheet.ListObjects.Count; chartCount = [int]$sheet.ChartObjects().Count }
            Release-ComObject $used
            Release-ComObject $sheet
          }
          $results += @{ op = $op; name = [string]$book.Name; path = [string]$book.FullName; sheets = $sheets }
        }
        "read_range" {
          $sheet = $book.Worksheets.Item([string]$operation.sheet)
          $range = $sheet.Range([string]$operation.range)
          $results += @{ op = $op; sheet = [string]$sheet.Name; data = Get-ExcelRangeSnapshot $range }
          Release-ComObject $range; Release-ComObject $sheet
        }
        "set_values" {
          $sheet = $book.Worksheets.Item([string]$operation.sheet)
          $start = $sheet.Range([string]$operation.startCell)
          $rowIndex = 0
          foreach ($row in $operation.values) {
            $columnIndex = 0
            foreach ($value in $row) {
              $cell = $start.Offset($rowIndex, $columnIndex)
              if ($null -eq $value) { $cell.ClearContents() }
              elseif ($value -is [string]) { $cell.Value2 = [string]$value }
              elseif ($value -is [bool]) { $cell.Value2 = [bool]$value }
              elseif ($value -is [int] -or $value -is [long] -or $value -is [double] -or $value -is [decimal]) { $cell.Value2 = [double]$value }
              else { $cell.Value2 = [string]$value }
              Release-ComObject $cell
              $columnIndex++
            }
            $rowIndex++
          }
          $results += @{ op = $op; rowsWritten = $rowIndex }
          Release-ComObject $start; Release-ComObject $sheet
        }
        "set_formula" {
          $sheet = $book.Worksheets.Item([string]$operation.sheet)
          $range = $sheet.Range([string]$operation.range)
          $range.Formula = [string]$operation.formula
          $results += @{ op = $op; range = [string]$range.Address($false, $false) }
          Release-ComObject $range; Release-ComObject $sheet
        }
        "format_range" {
          $sheet = $book.Worksheets.Item([string]$operation.sheet)
          $range = $sheet.Range([string]$operation.range)
          $format = $operation.format
          if ($null -ne $format.PSObject.Properties["bold"]) { $range.Font.Bold = [bool]$format.bold }
          if ($null -ne $format.PSObject.Properties["italic"]) { $range.Font.Italic = [bool]$format.italic }
          if ($null -ne $format.PSObject.Properties["fontSize"]) { $range.Font.Size = [double]$format.fontSize }
          if ($null -ne $format.PSObject.Properties["fontColor"]) { $range.Font.Color = Convert-HexColor ([string]$format.fontColor) }
          if ($null -ne $format.PSObject.Properties["fillColor"]) { $range.Interior.Color = Convert-HexColor ([string]$format.fillColor) }
          if ($null -ne $format.PSObject.Properties["numberFormat"]) { $range.NumberFormat = [string]$format.numberFormat }
          if ($null -ne $format.PSObject.Properties["wrapText"]) { $range.WrapText = [bool]$format.wrapText }
          if ($null -ne $format.PSObject.Properties["horizontalAlignment"]) {
            $alignments = @{ left = -4131; center = -4108; right = -4152; general = 1 }
            $range.HorizontalAlignment = $alignments[[string]$format.horizontalAlignment]
          }
          if ([bool](Get-Property $format "merge" $false)) { $range.Merge() }
          if ([bool](Get-Property $format "autoFitColumns" $false)) { [void]$range.Columns.AutoFit() }
          if ([bool](Get-Property $format "autoFitRows" $false)) { [void]$range.Rows.AutoFit() }
          $results += @{ op = $op; range = [string]$range.Address($false, $false) }
          Release-ComObject $range; Release-ComObject $sheet
        }
        "add_sheet" {
          $sheet = $book.Worksheets.Add()
          $sheet.Name = [string]$operation.name
          $results += @{ op = $op; name = [string]$sheet.Name }
          Release-ComObject $sheet
        }
        "define_name" {
          $sheet = $book.Worksheets.Item([string]$operation.sheet)
          $range = $sheet.Range([string]$operation.range)
          $sheetName = ([string]$sheet.Name).Replace("'", "''")
          $refersTo = "='$sheetName'!$($range.Address($true, $true))"
          $namedRange = $book.Names.Add([string]$operation.name, $refersTo)
          $results += @{ op = $op; name = [string]$namedRange.Name; refersTo = [string]$namedRange.RefersTo }
          Release-ComObject $namedRange; Release-ComObject $range; Release-ComObject $sheet
        }
        "set_conditional_format" {
          $sheet = $book.Worksheets.Item([string]$operation.sheet)
          $range = $sheet.Range([string]$operation.range)
          $operators = @{ between = 1; notBetween = 2; equal = 3; notEqual = 4; greater = 5; less = 6; greaterEqual = 7; lessEqual = 8 }
          $kind = [string]$operation.kind
          $type = if ($kind -eq "formula") { 2 } else { 1 }
          $operatorName = [string](Get-Property $operation "operator" "equal")
          $condition = $range.FormatConditions.Add($type, [int]$operators[$operatorName], [string]$operation.formula1, (Get-Property $operation "formula2" $null))
          if ($null -ne $operation.PSObject.Properties["fillColor"]) { $condition.Interior.Color = Convert-HexColor ([string]$operation.fillColor) }
          if ($null -ne $operation.PSObject.Properties["fontColor"]) { $condition.Font.Color = Convert-HexColor ([string]$operation.fontColor) }
          $results += @{ op = $op; range = [string]$range.Address($false, $false) }
          Release-ComObject $condition; Release-ComObject $range; Release-ComObject $sheet
        }
        "set_validation" {
          $sheet = $book.Worksheets.Item([string]$operation.sheet)
          $range = $sheet.Range([string]$operation.range)
          $types = @{ whole = 1; decimal = 2; list = 3; date = 4; time = 5; textLength = 6; custom = 7 }
          $operators = @{ between = 1; notBetween = 2; equal = 3; notEqual = 4; greater = 5; less = 6; greaterEqual = 7; lessEqual = 8 }
          $validation = $range.Validation
          $validation.Delete()
          $operatorName = [string](Get-Property $operation "operator" "between")
          $validation.Add([int]$types[[string]$operation.kind], 1, [int]$operators[$operatorName], [string]$operation.formula1, (Get-Property $operation "formula2" $null))
          if ($null -ne $operation.PSObject.Properties["errorMessage"]) {
            $validation.ErrorMessage = [string]$operation.errorMessage
            $validation.ShowError = $true
          }
          $results += @{ op = $op; range = [string]$range.Address($false, $false) }
          Release-ComObject $validation; Release-ComObject $range; Release-ComObject $sheet
        }
        "rename_sheet" {
          $sheet = $book.Worksheets.Item([string]$operation.sheet)
          $sheet.Name = [string]$operation.newName
          $results += @{ op = $op; name = [string]$sheet.Name }
          Release-ComObject $sheet
        }
        "delete_sheet" {
          $sheet = $book.Worksheets.Item([string]$operation.sheet)
          $sheet.Delete()
          $results += @{ op = $op; deleted = [string]$operation.sheet }
          Release-ComObject $sheet
        }
        "add_table" {
          $sheet = $book.Worksheets.Item([string]$operation.sheet)
          $range = $sheet.Range([string]$operation.range)
          $table = $sheet.ListObjects.Add(1, $range, $null, 1)
          $table.Name = [string]$operation.name
          if ($null -ne $operation.PSObject.Properties["style"]) { $table.TableStyle = [string]$operation.style }
          $results += @{ op = $op; name = [string]$table.Name }
          Release-ComObject $table; Release-ComObject $range; Release-ComObject $sheet
        }
        "add_chart" {
          $sheet = $book.Worksheets.Item([string]$operation.sheet)
          $source = $sheet.Range([string]$operation.sourceRange)
          $types = @{ column = 51; bar = 57; line = 4; pie = 5; doughnut = -4120; area = 1; scatter = -4169 }
          $shape = $sheet.Shapes.AddChart2(-1, $types[[string]$operation.chartType], [double]$operation.left, [double]$operation.top, [double]$operation.width, [double]$operation.height)
          $chart = $shape.Chart
          $chart.SetSourceData($source)
          if ($null -ne $operation.PSObject.Properties["title"]) { $chart.HasTitle = $true; $chart.ChartTitle.Text = [string]$operation.title }
          $results += @{ op = $op; name = [string]$shape.Name }
          Release-ComObject $chart; Release-ComObject $shape; Release-ComObject $source; Release-ComObject $sheet
        }
        "create_pivot" {
          $sourceSheet = $book.Worksheets.Item([string]$operation.sourceSheet)
          $source = $sourceSheet.Range([string]$operation.sourceRange)
          $destinationSheet = $book.Worksheets.Item([string]$operation.destinationSheet)
          $destination = $destinationSheet.Range([string]$operation.destinationCell)
          $sourceAddress = $source.Address($true, $true, 1, $true)
          $cache = $book.PivotCaches().Create(1, $sourceAddress)
          $pivot = $cache.CreatePivotTable($destination, [string]$operation.name)
          foreach ($fieldName in @(Get-Property $operation "rows" @())) {
            $field = $pivot.PivotFields([string]$fieldName)
            $field.Orientation = 1
            Release-ComObject $field
          }
          foreach ($fieldName in @(Get-Property $operation "columns" @())) {
            $field = $pivot.PivotFields([string]$fieldName)
            $field.Orientation = 2
            Release-ComObject $field
          }
          foreach ($fieldName in @(Get-Property $operation "filters" @())) {
            $field = $pivot.PivotFields([string]$fieldName)
            $field.Orientation = 3
            Release-ComObject $field
          }
          $functions = @{ sum = -4157; count = -4112; average = -4106; max = -4136; min = -4139 }
          foreach ($valueDefinition in @(Get-Property $operation "values" @())) {
            $field = $pivot.PivotFields([string]$valueDefinition.field)
            $caption = [string](Get-Property $valueDefinition "caption" "Sum of $($valueDefinition.field)")
            $dataField = $pivot.AddDataField($field, $caption, $functions[[string](Get-Property $valueDefinition "aggregation" "sum")])
            if ($null -ne $valueDefinition.PSObject.Properties["numberFormat"]) { $dataField.NumberFormat = [string]$valueDefinition.numberFormat }
            Release-ComObject $dataField; Release-ComObject $field
          }
          $results += @{ op = $op; name = [string]$pivot.Name; destination = [string]$destination.Address($false, $false) }
          Release-ComObject $pivot; Release-ComObject $cache; Release-ComObject $destination; Release-ComObject $destinationSheet; Release-ComObject $source; Release-ComObject $sourceSheet
        }
        "sort" {
          $sheet = $book.Worksheets.Item([string]$operation.sheet)
          $range = $sheet.Range([string]$operation.range)
          $key = $sheet.Range([string]$operation.key)
          $order = if ([bool](Get-Property $operation "ascending" $true)) { 1 } else { 2 }
          [void]$range.Sort($key, $order, $null, $null, $null, $null, $null, 1)
          $results += @{ op = $op; sorted = $true }
          Release-ComObject $key; Release-ComObject $range; Release-ComObject $sheet
        }
        "autofilter" {
          $sheet = $book.Worksheets.Item([string]$operation.sheet)
          $range = $sheet.Range([string]$operation.range)
          [void]$range.AutoFilter([int]$operation.field, (Get-Property $operation "criteria" $null))
          $results += @{ op = $op; filtered = $true }
          Release-ComObject $range; Release-ComObject $sheet
        }
        "recalculate" {
          $app.CalculateFullRebuild()
          $results += @{ op = $op; recalculated = $true }
        }
        "save" {
          $outputPath = [string](Get-Property $operation "outputPath" "")
          if ([string]::IsNullOrWhiteSpace($outputPath)) { $book.Save() } else { $book.SaveAs($outputPath, 51) }
          $results += @{ op = $op; path = [string]$book.FullName }
        }
        "export_pdf" {
          $book.ExportAsFixedFormat(0, [string]$operation.outputPath)
          $results += @{ op = $op; path = [string]$operation.outputPath }
        }
        default { throw "Unsupported Excel operation: $op" }
      }
    }
    return @{ application = "excel"; results = $results }
  } finally {
    if ($session.openedHere) { try { $book.Close($false) } catch {} }
    if ($session.ownedApp) { try { $app.Quit() } catch {} }
    Release-ComObject $book; Release-ComObject $app
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
  }
}

function Get-WordSession($Target) {
  $attach = [bool](Get-Property $Target "active" $false)
  $app = if ($attach) { Get-ActiveComObject "Word.Application" } else { $null }
  $ownedApp = $null -eq $app
  if ($ownedApp) { $app = New-Object -ComObject Word.Application }
  if ($ownedApp) { $app.Visible = [bool](Get-Property $Target "visible" $false) }
  if ($ownedApp) { $app.DisplayAlerts = 0 }

  $document = $null
  $openedHere = $false
  $targetPath = [string](Get-Property $Target "path" "")
  if ([bool](Get-Property $Target "create" $false)) {
    $document = $app.Documents.Add()
    $openedHere = $true
  } elseif (-not [string]::IsNullOrWhiteSpace($targetPath)) {
    foreach ($candidate in $app.Documents) {
      if ([string]::Equals([string]$candidate.FullName, $targetPath, [StringComparison]::OrdinalIgnoreCase)) {
        $document = $candidate
        break
      }
      Release-ComObject $candidate
    }
    if ($null -eq $document) { $document = $app.Documents.Open($targetPath); $openedHere = $true }
  } else {
    $document = $app.ActiveDocument
  }
  if ($null -eq $document) { throw "No Word document is available. Provide path, create=true, or active=true." }
  return @{ app = $app; document = $document; ownedApp = $ownedApp; openedHere = $openedHere }
}

function Invoke-WordBatch($Request) {
  $session = Get-WordSession $Request.target
  $app = $session.app
  $document = $session.document
  $results = @()
  try {
    foreach ($operation in $Request.operations) {
      $op = [string]$operation.op
      switch ($op) {
        "inspect" {
          $headings = @()
          foreach ($paragraph in $document.Paragraphs) {
            $text = ([string]$paragraph.Range.Text).Trim([char]13, [char]7, [char]32)
            $styleName = [string]$paragraph.Range.Style
            try { $styleName = [string]$paragraph.Range.Style.NameLocal } catch {}
            if ($styleName -match "Heading|Title") { $headings += @{ style = $styleName; text = $text } }
            Release-ComObject $paragraph
          }
          $results += @{ op = $op; name = [string]$document.Name; path = [string]$document.FullName; words = [int]$document.Words.Count; paragraphs = [int]$document.Paragraphs.Count; tables = [int]$document.Tables.Count; sections = [int]$document.Sections.Count; tablesOfContents = [int]$document.TablesOfContents.Count; comments = [int]$document.Comments.Count; headings = $headings; trackRevisions = [bool]$document.TrackRevisions }
        }
        "read_text" {
          $text = [string]$document.Content.Text
          $maxCharacters = [int](Get-Property $operation "maxCharacters" 100000)
          if ($text.Length -gt $maxCharacters) { $text = $text.Substring(0, $maxCharacters) }
          $results += @{ op = $op; text = $text; truncated = $document.Content.Text.Length -gt $maxCharacters }
        }
        "apply_style" {
          $paragraph = $document.Paragraphs.Item([int]$operation.paragraphIndex)
          $paragraph.Range.Style = [string]$operation.style
          $results += @{ op = $op; paragraphIndex = [int]$operation.paragraphIndex; style = [string]$operation.style }
          Release-ComObject $paragraph
        }
        "insert_section" {
          $range = $document.Content
          $range.Collapse(0)
          $breakType = [string](Get-Property $operation "breakType" "nextPage")
          $breakCode = if ($breakType -eq "continuous") { 3 } else { 2 }
          $range.InsertBreak($breakCode)
          $section = $document.Sections.Item($document.Sections.Count)
          if ($null -ne $operation.PSObject.Properties["orientation"]) {
            $section.PageSetup.Orientation = if ([string]$operation.orientation -eq "landscape") { 1 } else { 0 }
          }
          $results += @{ op = $op; sections = [int]$document.Sections.Count }
          Release-ComObject $section; Release-ComObject $range
        }
        "create_toc" {
          $range = if ([string](Get-Property $operation "position" "start") -eq "end") { $document.Content } else { $document.Range(0, 0) }
          if ([string](Get-Property $operation "position" "start") -eq "end") { $range.Collapse(0) }
          $toc = $document.TablesOfContents.Add($range, $true, 1, [int](Get-Property $operation "maxLevel" 3))
          $toc.Update()
          $results += @{ op = $op; tablesOfContents = [int]$document.TablesOfContents.Count }
          Release-ComObject $toc; Release-ComObject $range
        }
        "update_toc" {
          foreach ($toc in $document.TablesOfContents) { $toc.Update(); Release-ComObject $toc }
          $results += @{ op = $op; tablesOfContents = [int]$document.TablesOfContents.Count }
        }
        "add_comment" {
          $paragraph = $document.Paragraphs.Item([int]$operation.paragraphIndex)
          $comment = $document.Comments.Add($paragraph.Range, [string]$operation.text)
          $results += @{ op = $op; comments = [int]$document.Comments.Count }
          Release-ComObject $comment; Release-ComObject $paragraph
        }
        "append_text" {
          $range = $document.Content
          $range.Collapse(0)
          $range.InsertAfter([string]$operation.text)
          if ([bool](Get-Property $operation "newParagraph" $true)) { $range.InsertParagraphAfter() }
          if ($null -ne $operation.PSObject.Properties["style"]) { $range.Style = [string]$operation.style }
          $results += @{ op = $op; appended = $true }
          Release-ComObject $range
        }
        "add_heading" {
          $range = $document.Content
          $range.Collapse(0)
          $range.InsertAfter([string]$operation.text)
          $range.Style = "Heading $([int](Get-Property $operation 'level' 1))"
          $range.InsertParagraphAfter()
          $results += @{ op = $op; added = $true }
          Release-ComObject $range
        }
        "replace_text" {
          $range = $document.Content
          $find = $range.Find
          $find.ClearFormatting()
          $find.Replacement.ClearFormatting()
          $matchCase = [bool](Get-Property $operation "matchCase" $true)
          $replaced = $find.Execute([string]$operation.find, $matchCase, $false, $false, $false, $false, $true, 1, $false, [string]$operation.replacement, 2)
          $results += @{ op = $op; replaced = [bool]$replaced }
          Release-ComObject $find; Release-ComObject $range
        }
        "add_table" {
          $rows = @($operation.rows)
          $rowCount = $rows.Count
          $columnCount = @($rows[0]).Count
          $range = $document.Content
          $range.Collapse(0)
          $table = $document.Tables.Add($range, $rowCount, $columnCount)
          for ($row = 1; $row -le $rowCount; $row++) {
            for ($column = 1; $column -le $columnCount; $column++) {
              $cell = $table.Cell($row, $column)
              $cell.Range.Text = [string]$rows[$row - 1][$column - 1]
              if ([bool](Get-Property $operation "headerRow" $false) -and $row -eq 1) { $cell.Range.Bold = 1 }
              Release-ComObject $cell
            }
          }
          if ($null -ne $operation.PSObject.Properties["style"]) { $table.Style = [string]$operation.style }
          [void]$table.AutoFitBehavior(1)
          $results += @{ op = $op; rows = $rowCount; columns = $columnCount }
          Release-ComObject $table; Release-ComObject $range
        }
        "set_track_changes" {
          $document.TrackRevisions = [bool]$operation.enabled
          $results += @{ op = $op; enabled = [bool]$document.TrackRevisions }
        }
        "save" {
          $outputPath = [string](Get-Property $operation "outputPath" "")
          if ([string]::IsNullOrWhiteSpace($outputPath)) { $document.Save() } else { $document.SaveAs2($outputPath, 16) }
          $results += @{ op = $op; path = [string]$document.FullName }
        }
        "export_pdf" {
          $document.ExportAsFixedFormat([string]$operation.outputPath, 17)
          $results += @{ op = $op; path = [string]$operation.outputPath }
        }
        default { throw "Unsupported Word operation: $op" }
      }
    }
    return @{ application = "word"; results = $results }
  } finally {
    if ($session.openedHere) { try { $document.Close(0) } catch {} }
    if ($session.ownedApp) { try { $app.Quit() } catch {} }
    Release-ComObject $document; Release-ComObject $app
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
  }
}

function Get-PowerPointSession($Target) {
  $attach = [bool](Get-Property $Target "active" $false)
  $runningPowerPoint = @(Get-Process -Name "POWERPNT" -ErrorAction SilentlyContinue)
  for ($attempt = 0; -not $attach -and $runningPowerPoint.Count -gt 0 -and $attempt -lt 30; $attempt++) {
    Start-Sleep -Milliseconds 500
    $runningPowerPoint = @(Get-Process -Name "POWERPNT" -ErrorAction SilentlyContinue)
  }
  if (-not $attach -and $runningPowerPoint.Count -gt 0) {
    throw "PowerPoint is already running. Background native work would share the user's PowerPoint session. Close PowerPoint or explicitly opt in with target.active=true."
  }
  $app = if ($attach) { Get-ActiveComObject "PowerPoint.Application" } else { New-Object -ComObject PowerPoint.Application }
  if ($null -eq $app) { throw "No running PowerPoint application is available for target.active=true." }
  $ownedApp = -not $attach
  $showWindow = [bool](Get-Property $Target "visible" $false)
  $withWindow = if ($showWindow) { -1 } else { 0 }
  if ($ownedApp -and $showWindow) { $app.Visible = -1 }

  $presentation = $null
  $openedHere = $false
  $targetPath = [string](Get-Property $Target "path" "")
  if ([bool](Get-Property $Target "create" $false)) {
    $presentation = $app.Presentations.Add($withWindow)
    $openedHere = $true
  } elseif (-not [string]::IsNullOrWhiteSpace($targetPath)) {
    foreach ($candidate in $app.Presentations) {
      if ([string]::Equals([string]$candidate.FullName, $targetPath, [StringComparison]::OrdinalIgnoreCase)) {
        $presentation = $candidate
        break
      }
      Release-ComObject $candidate
    }
    if ($null -eq $presentation) { $presentation = $app.Presentations.Open($targetPath, 0, 0, $withWindow); $openedHere = $true }
  } else {
    $presentation = $app.ActivePresentation
  }
  if ($null -eq $presentation) { throw "No PowerPoint presentation is available. Provide path, create=true, or active=true." }
  return @{ app = $app; document = $presentation; ownedApp = $ownedApp; openedHere = $openedHere }
}

function Invoke-PowerPointBatch($Request) {
  $session = Get-PowerPointSession $Request.target
  $app = $session.app
  $presentation = $session.document
  $results = @()
  try {
    foreach ($operation in $Request.operations) {
      $op = [string]$operation.op
      switch ($op) {
        "inspect" {
          $slides = @()
          $designs = @()
          foreach ($design in $presentation.Designs) {
            $layouts = @()
            foreach ($layout in $design.SlideMaster.CustomLayouts) {
              $placeholders = @()
              try {
                for ($placeholderIndex = 1; $placeholderIndex -le [int]$layout.Shapes.Placeholders.Count; $placeholderIndex++) {
                  $placeholder = $layout.Shapes.Placeholders.Item($placeholderIndex)
                  $placeholders += @{
                    index = $placeholderIndex
                    name = [string]$placeholder.Name
                    type = [int]$placeholder.PlaceholderFormat.Type
                    hasTextFrame = [bool]($placeholder.HasTextFrame -eq -1)
                    left = [double]$placeholder.Left
                    top = [double]$placeholder.Top
                    width = [double]$placeholder.Width
                    height = [double]$placeholder.Height
                  }
                  Release-ComObject $placeholder
                }
              } catch { }
              $layouts += @{ index = [int]$layout.Index; name = [string]$layout.Name; placeholders = $placeholders }
              Release-ComObject $layout
            }
            $designs += @{ index = [int]$design.Index; name = [string]$design.Name; layouts = $layouts }
            Release-ComObject $design
          }
          foreach ($slide in $presentation.Slides) {
            $texts = @()
            foreach ($shape in $slide.Shapes) {
              if ($shape.HasTextFrame -and $shape.TextFrame.HasText) { $texts += [string]$shape.TextFrame.TextRange.Text }
              Release-ComObject $shape
            }
            $slides += @{ number = [int]$slide.SlideIndex; name = [string]$slide.Name; shapeCount = [int]$slide.Shapes.Count; text = $texts }
            Release-ComObject $slide
          }
          $results += @{ op = $op; name = [string]$presentation.Name; path = [string]$presentation.FullName; width = [double]$presentation.PageSetup.SlideWidth; height = [double]$presentation.PageSetup.SlideHeight; presentationWindowCount = [int]$presentation.Windows.Count; designs = $designs; slides = $slides }
        }
        "list_shapes" {
          $slide = $presentation.Slides.Item([int]$operation.slide)
          $shapes = @()
          $shapeIndex = 0
          foreach ($shape in $slide.Shapes) {
            $shapeIndex++
            $shapeText = ""
            if ($shape.HasTextFrame -and $shape.TextFrame.HasText) { $shapeText = [string]$shape.TextFrame.TextRange.Text }
            $shapes += @{ index = $shapeIndex; name = [string]$shape.Name; type = [int]$shape.Type; text = $shapeText }
            Release-ComObject $shape
          }
          $results += @{ op = $op; slide = [int]$slide.SlideIndex; shapes = $shapes }
          Release-ComObject $slide
        }
        "list_placeholders" {
          $slide = $presentation.Slides.Item([int]$operation.slide)
          $placeholders = @()
          for ($i = 1; $i -le [int]$slide.Shapes.Placeholders.Count; $i++) {
            $placeholder = $slide.Shapes.Placeholders.Item($i)
            $placeholders += @{ index = $i; name = [string]$placeholder.Name; type = [int]$placeholder.PlaceholderFormat.Type; hasTextFrame = [bool]($placeholder.HasTextFrame -eq -1); text = if ($placeholder.HasTextFrame -and $placeholder.TextFrame.HasText) { [string]$placeholder.TextFrame.TextRange.Text } else { "" } }
            Release-ComObject $placeholder
          }
          $results += @{ op = $op; slide = [int]$slide.SlideIndex; placeholders = $placeholders }
          Release-ComObject $slide
        }
        "set_placeholder_text" {
          $slide = $presentation.Slides.Item([int]$operation.slide)
          $placeholder = $slide.Shapes.Placeholders.Item([int]$operation.placeholderIndex)
          if (-not $placeholder.HasTextFrame) { throw "Placeholder does not support text: $($operation.placeholderIndex)" }
          $placeholder.TextFrame.TextRange.Text = [string]$operation.text
          $results += @{ op = $op; slide = [int]$slide.SlideIndex; placeholderIndex = [int]$operation.placeholderIndex }
          Release-ComObject $placeholder; Release-ComObject $slide
        }
        "list_animations" {
          $slide = $presentation.Slides.Item([int]$operation.slide)
          $animations = @()
          $sequence = $null
          try { $sequence = $slide.TimeLine.MainSequence } catch { }
          if ($null -ne $sequence) {
            for ($i = 1; $i -le [int]$sequence.Count; $i++) {
              $effect = $sequence.Item($i)
              $shape = $effect.Shape
              $timing = $effect.Timing
              $animations += @{
                index = $i
                shape = [string]$shape.Name
                effectId = [int]$effect.EffectType
                exit = [bool]($effect.Exit -eq -1)
                trigger = [int]$timing.TriggerType
                durationSeconds = [double]$timing.Duration
                delaySeconds = [double]$timing.TriggerDelayTime
              }
              Release-ComObject $timing; Release-ComObject $shape; Release-ComObject $effect
            }
            Release-ComObject $sequence
          }
          $transition = $slide.SlideShowTransition
          $transitionInfo = @{
            effectId = [int]$transition.EntryEffect
            advanceOnClick = [bool]($transition.AdvanceOnClick -eq -1)
            advanceOnTime = [bool]($transition.AdvanceOnTime -eq -1)
            advanceAfterSeconds = [double]$transition.AdvanceTime
          }
          $results += @{ op = $op; slide = [int]$slide.SlideIndex; animations = $animations; transition = $transitionInfo }
          Release-ComObject $transition
          Release-ComObject $slide
        }
        "add_slide" {
          $layout = [int](Get-Property $operation "layout" 12)
          $slide = $presentation.Slides.Add($presentation.Slides.Count + 1, $layout)
          $results += @{ op = $op; slide = [int]$slide.SlideIndex }
          Release-ComObject $slide
        }
        "add_slide_from_layout" {
          $design = $presentation.Designs.Item([int]$operation.designIndex)
          $layout = $design.SlideMaster.CustomLayouts.Item([int]$operation.layoutIndex)
          $slide = $presentation.Slides.AddSlide($presentation.Slides.Count + 1, $layout)
          $results += @{ op = $op; slide = [int]$slide.SlideIndex; layout = [string]$layout.Name; design = [string]$design.Name }
          Release-ComObject $slide; Release-ComObject $layout; Release-ComObject $design
        }
        "duplicate_slide" {
          $source = $presentation.Slides.Item([int]$operation.slide)
          $duplicated = $source.Duplicate()
          $slide = $duplicated.Item(1)
          if ($null -ne $operation.PSObject.Properties["toIndex"]) { $slide.MoveTo([int]$operation.toIndex) }
          $results += @{ op = $op; slide = [int]$slide.SlideIndex }
          Release-ComObject $slide; Release-ComObject $duplicated; Release-ComObject $source
        }
        "move_slide" {
          $slide = $presentation.Slides.Item([int]$operation.slide)
          $slide.MoveTo([int]$operation.toIndex)
          $results += @{ op = $op; slide = [int]$slide.SlideIndex }
          Release-ComObject $slide
        }
        "add_text" {
          $slide = $presentation.Slides.Item([int]$operation.slide)
          $shape = $slide.Shapes.AddTextbox(1, [double]$operation.left, [double]$operation.top, [double]$operation.width, [double]$operation.height)
          $shape.TextFrame.TextRange.Text = [string]$operation.text
          $shape.TextFrame.TextRange.Font.Size = [double](Get-Property $operation "fontSize" 18)
          if ($null -ne $operation.PSObject.Properties["fontFace"]) { $shape.TextFrame.TextRange.Font.Name = [string]$operation.fontFace }
          if ($null -ne $operation.PSObject.Properties["color"]) { $shape.TextFrame.TextRange.Font.Color.RGB = Convert-HexColor ([string]$operation.color) }
          if ([bool](Get-Property $operation "bold" $false)) { $shape.TextFrame.TextRange.Font.Bold = -1 }
          $results += @{ op = $op; slide = [int]$slide.SlideIndex; shape = [string]$shape.Name }
          Release-ComObject $shape; Release-ComObject $slide
        }
        "add_shape" {
          $slide = $presentation.Slides.Item([int]$operation.slide)
          $shapeTypes = @{ rectangle = 1; roundRectangle = 5; ellipse = 9; triangle = 7; diamond = 4; chevron = 52 }
          $shape = $slide.Shapes.AddShape($shapeTypes[[string]$operation.shape], [double]$operation.left, [double]$operation.top, [double]$operation.width, [double]$operation.height)
          if ($null -ne $operation.PSObject.Properties["fillColor"]) { $shape.Fill.ForeColor.RGB = Convert-HexColor ([string]$operation.fillColor) }
          if ($null -ne $operation.PSObject.Properties["lineColor"]) { $shape.Line.ForeColor.RGB = Convert-HexColor ([string]$operation.lineColor) }
          if ($null -ne $operation.PSObject.Properties["text"]) {
            if (-not $shape.HasTextFrame) { throw "Shape does not support text: $($operation.name)" }
            $shape.TextFrame.TextRange.Text = [string]$operation.text
          }
          $results += @{ op = $op; slide = [int]$slide.SlideIndex; shape = [string]$shape.Name }
          Release-ComObject $shape; Release-ComObject $slide
        }
        "add_picture" {
          $slide = $presentation.Slides.Item([int]$operation.slide)
          $shape = $slide.Shapes.AddPicture([string]$operation.path, 0, -1, [double]$operation.left, [double]$operation.top, [double]$operation.width, [double]$operation.height)
          if ($null -ne $operation.PSObject.Properties["altText"]) { $shape.AlternativeText = [string]$operation.altText }
          $results += @{ op = $op; slide = [int]$slide.SlideIndex; shape = [string]$shape.Name }
          Release-ComObject $shape; Release-ComObject $slide
        }
        "add_media" {
          $slide = $presentation.Slides.Item([int]$operation.slide)
          $shape = $slide.Shapes.AddMediaObject2([string]$operation.path, 0, -1, [double]$operation.left, [double]$operation.top, [double]$operation.width, [double]$operation.height)
          $results += @{ op = $op; slide = [int]$slide.SlideIndex; shape = [string]$shape.Name }
          Release-ComObject $shape; Release-ComObject $slide
        }
        "rename_shape" {
          $slide = $presentation.Slides.Item([int]$operation.slide)
          $selector = if ($operation.shape -is [string]) { [string]$operation.shape } else { [int]$operation.shape }
          $shape = $slide.Shapes.Item($selector)
          $oldName = [string]$shape.Name
          $shape.Name = [string]$operation.name
          $results += @{ op = $op; slide = [int]$slide.SlideIndex; oldName = $oldName; name = [string]$shape.Name }
          Release-ComObject $shape; Release-ComObject $slide
        }
        "group_shapes" {
          $slide = $presentation.Slides.Item([int]$operation.slide)
          [object[]]$selectors = @()
          foreach ($item in $operation.shapes) { $selectors += if ($item -is [string]) { [string]$item } else { [int]$item } }
          $range = $slide.Shapes.Range($selectors)
          $shape = $range.Group()
          if ($null -ne $operation.PSObject.Properties["name"]) { $shape.Name = [string]$operation.name }
          $results += @{ op = $op; slide = [int]$slide.SlideIndex; shape = [string]$shape.Name; itemCount = [int]$shape.GroupItems.Count }
          Release-ComObject $shape; Release-ComObject $range; Release-ComObject $slide
        }
        "ungroup_shape" {
          $slide = $presentation.Slides.Item([int]$operation.slide)
          $selector = if ($operation.shape -is [string]) { [string]$operation.shape } else { [int]$operation.shape }
          $shape = $slide.Shapes.Item($selector)
          $range = $shape.Ungroup()
          $names = @()
          foreach ($item in $range) { $names += [string]$item.Name; Release-ComObject $item }
          $results += @{ op = $op; slide = [int]$slide.SlideIndex; shapes = $names }
          Release-ComObject $range; Release-ComObject $shape; Release-ComObject $slide
        }
        "set_z_order" {
          $actions = @{ bringToFront = 0; sendToBack = 1; bringForward = 2; sendBackward = 3 }
          $name = [string]$operation.action
          if (-not $actions.ContainsKey($name)) { throw "Unsupported z-order action: $name" }
          $slide = $presentation.Slides.Item([int]$operation.slide)
          $selector = if ($operation.shape -is [string]) { [string]$operation.shape } else { [int]$operation.shape }
          $shape = $slide.Shapes.Item($selector)
          $shape.ZOrder([int]$actions[$name])
          $results += @{ op = $op; slide = [int]$slide.SlideIndex; shape = [string]$shape.Name; action = $name; zOrderPosition = [int]$shape.ZOrderPosition }
          Release-ComObject $shape; Release-ComObject $slide
        }
        "update_shape" {
          $slide = $presentation.Slides.Item([int]$operation.slide)
          $shape = $slide.Shapes.Item([string]$operation.name)
          if ($null -ne $operation.PSObject.Properties["left"]) { $shape.Left = [double]$operation.left }
          if ($null -ne $operation.PSObject.Properties["top"]) { $shape.Top = [double]$operation.top }
          if ($null -ne $operation.PSObject.Properties["width"]) { $shape.Width = [double]$operation.width }
          if ($null -ne $operation.PSObject.Properties["height"]) { $shape.Height = [double]$operation.height }
          if ($null -ne $operation.PSObject.Properties["text"]) {
            if (-not $shape.HasTextFrame) { throw "Shape does not support text: $($operation.name)" }
            $shape.TextFrame.TextRange.Text = [string]$operation.text
          }
          if ($null -ne $operation.PSObject.Properties["fillColor"]) { $shape.Fill.ForeColor.RGB = Convert-HexColor ([string]$operation.fillColor) }
          if ($null -ne $operation.PSObject.Properties["lineColor"]) { $shape.Line.ForeColor.RGB = Convert-HexColor ([string]$operation.lineColor) }
          if ($null -ne $operation.PSObject.Properties["fontSize"]) { $shape.TextFrame.TextRange.Font.Size = [double]$operation.fontSize }
          if ($null -ne $operation.PSObject.Properties["fontColor"]) { $shape.TextFrame.TextRange.Font.Color.RGB = Convert-HexColor ([string]$operation.fontColor) }
          if ($null -ne $operation.PSObject.Properties["altText"]) { $shape.AlternativeText = [string]$operation.altText }
          if ($null -ne $operation.PSObject.Properties["rotation"]) { $shape.Rotation = [double]$operation.rotation }
          if ($null -ne $operation.PSObject.Properties["fillTransparency"]) { $shape.Fill.Transparency = [double]$operation.fillTransparency }
          if ($null -ne $operation.PSObject.Properties["lineTransparency"]) { $shape.Line.Transparency = [double]$operation.lineTransparency }
          $results += @{ op = $op; slide = [int]$slide.SlideIndex; shape = [string]$shape.Name }
          Release-ComObject $shape; Release-ComObject $slide
        }
        "set_speaker_notes" {
          $slide = $presentation.Slides.Item([int]$operation.slide)
          $notesShape = $slide.NotesPage.Shapes.Placeholders.Item(2)
          $notesShape.TextFrame.TextRange.Text = [string]$operation.text
          $results += @{ op = $op; slide = [int]$slide.SlideIndex }
          Release-ComObject $notesShape; Release-ComObject $slide
        }
        "add_animation" {
          $phase = [string](Get-Property $operation "phase" "entrance")
          $name = [string]$operation.effect
          $effects = @{
            entrance = @{ appear = 1; fade = 10; fly = 2; wipe = 22; zoom = 23 }
            emphasis = @{ spin = 61; growShrink = 59 }
            exit = @{ fade = 10; fly = 2; wipe = 22; zoom = 23 }
          }
          if (-not $effects.ContainsKey($phase) -or -not $effects[$phase].ContainsKey($name)) {
            throw "Animation effect '$name' is not supported for phase '$phase'."
          }
          $triggers = @{ onClick = 1; withPrevious = 2; afterPrevious = 3 }
          $triggerName = [string](Get-Property $operation "trigger" "onClick")
          if (-not $triggers.ContainsKey($triggerName)) { throw "Unsupported animation trigger: $triggerName" }
          $slide = $presentation.Slides.Item([int]$operation.slide)
          $selector = if ($operation.shape -is [string]) { [string]$operation.shape } else { [int]$operation.shape }
          $shape = $slide.Shapes.Item($selector)
          $position = [int](Get-Property $operation "position" -1)
          $effect = $slide.TimeLine.MainSequence.AddEffect($shape, [int]$effects[$phase][$name], 0, [int]$triggers[$triggerName], $position)
          if ($phase -eq "exit") { $effect.Exit = -1 }
          $timing = $effect.Timing
          $timing.Duration = [double](Get-Property $operation "durationSeconds" 0.6)
          $timing.TriggerDelayTime = [double](Get-Property $operation "delaySeconds" 0)
          if ($null -ne $operation.PSObject.Properties["repeatCount"]) { $timing.RepeatCount = [int]$operation.repeatCount }
          if ($null -ne $operation.PSObject.Properties["autoReverse"]) { $timing.AutoReverse = if ([bool]$operation.autoReverse) { -1 } else { 0 } }
          $results += @{ op = $op; slide = [int]$slide.SlideIndex; shape = [string]$shape.Name; animationIndex = [int]$effect.Index; phase = $phase; effect = $name; trigger = $triggerName }
          Release-ComObject $timing; Release-ComObject $effect; Release-ComObject $shape; Release-ComObject $slide
        }
        "update_animation" {
          $slide = $presentation.Slides.Item([int]$operation.slide)
          $effect = $slide.TimeLine.MainSequence.Item([int]$operation.animationIndex)
          $timing = $effect.Timing
          if ($null -ne $operation.PSObject.Properties["trigger"]) {
            $triggers = @{ onClick = 1; withPrevious = 2; afterPrevious = 3 }
            $timing.TriggerType = [int]$triggers[[string]$operation.trigger]
          }
          if ($null -ne $operation.PSObject.Properties["durationSeconds"]) { $timing.Duration = [double]$operation.durationSeconds }
          if ($null -ne $operation.PSObject.Properties["delaySeconds"]) { $timing.TriggerDelayTime = [double]$operation.delaySeconds }
          if ($null -ne $operation.PSObject.Properties["repeatCount"]) { $timing.RepeatCount = [int]$operation.repeatCount }
          if ($null -ne $operation.PSObject.Properties["autoReverse"]) { $timing.AutoReverse = if ([bool]$operation.autoReverse) { -1 } else { 0 } }
          $results += @{ op = $op; slide = [int]$slide.SlideIndex; animationIndex = [int]$effect.Index; durationSeconds = [double]$timing.Duration; delaySeconds = [double]$timing.TriggerDelayTime }
          Release-ComObject $timing; Release-ComObject $effect; Release-ComObject $slide
        }
        "move_animation" {
          $slide = $presentation.Slides.Item([int]$operation.slide)
          $sequence = $slide.TimeLine.MainSequence
          $toIndex = [int]$operation.toIndex
          if ($toIndex -gt [int]$sequence.Count) { throw "Animation destination index exceeds sequence length: $toIndex" }
          $effect = $sequence.Item([int]$operation.animationIndex)
          $effect.MoveTo($toIndex)
          $results += @{ op = $op; slide = [int]$slide.SlideIndex; animationIndex = [int]$effect.Index }
          Release-ComObject $effect; Release-ComObject $sequence; Release-ComObject $slide
        }
        "delete_animation" {
          $slide = $presentation.Slides.Item([int]$operation.slide)
          $effect = $slide.TimeLine.MainSequence.Item([int]$operation.animationIndex)
          $effect.Delete()
          $results += @{ op = $op; slide = [int]$slide.SlideIndex; deleted = [int]$operation.animationIndex }
          Release-ComObject $effect; Release-ComObject $slide
        }
        "clear_animations" {
          $slide = $presentation.Slides.Item([int]$operation.slide)
          $sequence = $null
          try { $sequence = $slide.TimeLine.MainSequence } catch { }
          $deleted = 0
          if ($null -ne $sequence) {
            for ($i = [int]$sequence.Count; $i -ge 1; $i--) {
              $effect = $sequence.Item($i)
              $effect.Delete()
              $deleted++
              Release-ComObject $effect
            }
            Release-ComObject $sequence
          }
          $results += @{ op = $op; slide = [int]$slide.SlideIndex; deleted = $deleted }
          Release-ComObject $slide
        }
        "set_transition" {
          $effects = @{ none = 0; cut = 257; fade = 1793; pushLeft = 3853; pushRight = 3854; wipeLeft = 2817; wipeRight = 2819; zoomIn = 3345; morph = 3954; morphWords = 3955; morphCharacters = 3956 }
          $name = [string]$operation.effect
          if (-not $effects.ContainsKey($name)) { throw "Unsupported slide transition: $name" }
          $slide = $presentation.Slides.Item([int]$operation.slide)
          $transition = $slide.SlideShowTransition
          $transition.EntryEffect = [int]$effects[$name]
          if ($null -ne $operation.PSObject.Properties["durationSeconds"]) { $transition.Duration = [double]$operation.durationSeconds }
          $transition.AdvanceOnClick = if ([bool](Get-Property $operation "advanceOnClick" $true)) { -1 } else { 0 }
          if ($null -ne $operation.PSObject.Properties["advanceAfterSeconds"]) {
            $transition.AdvanceOnTime = -1
            $transition.AdvanceTime = [double]$operation.advanceAfterSeconds
            $presentation.SlideShowSettings.AdvanceMode = 2
          } else {
            $transition.AdvanceOnTime = 0
          }
          $results += @{ op = $op; slide = [int]$slide.SlideIndex; effect = $name; advanceOnClick = [bool](Get-Property $operation "advanceOnClick" $true); advanceAfterSeconds = Get-Property $operation "advanceAfterSeconds" $null }
          Release-ComObject $transition; Release-ComObject $slide
        }
        "replace_text" {
          $count = 0
          foreach ($slide in $presentation.Slides) {
            foreach ($shape in $slide.Shapes) {
              if ($shape.HasTextFrame -and $shape.TextFrame.HasText) {
                $current = [string]$shape.TextFrame.TextRange.Text
                $comparison = if ([bool](Get-Property $operation "matchCase" $true)) { [StringComparison]::Ordinal } else { [StringComparison]::OrdinalIgnoreCase }
                $index = $current.IndexOf([string]$operation.find, $comparison)
                while ($index -ge 0) {
                  $current = $current.Remove($index, ([string]$operation.find).Length).Insert($index, [string]$operation.replacement)
                  $count++
                  $index = $current.IndexOf([string]$operation.find, $index + ([string]$operation.replacement).Length, $comparison)
                }
                $shape.TextFrame.TextRange.Text = $current
              }
              Release-ComObject $shape
            }
            Release-ComObject $slide
          }
          $results += @{ op = $op; replacements = $count }
        }
        "delete_slide" {
          $presentation.Slides.Item([int]$operation.slide).Delete()
          $results += @{ op = $op; deleted = [int]$operation.slide }
        }
        "save" {
          $outputPath = [string](Get-Property $operation "outputPath" "")
          if ([string]::IsNullOrWhiteSpace($outputPath)) { $presentation.Save() } else { $presentation.SaveAs($outputPath, 24) }
          $results += @{ op = $op; path = [string]$presentation.FullName }
        }
        "export_pdf" {
          $presentation.SaveAs([string]$operation.outputPath, 32)
          $results += @{ op = $op; path = [string]$operation.outputPath }
        }
        "render_slides" {
          $presentation.Export([string]$operation.outputDirectory, "PNG", [int](Get-Property $operation "width" 1600), [int](Get-Property $operation "height" 900))
          $results += @{ op = $op; outputDirectory = [string]$operation.outputDirectory; slideCount = [int]$presentation.Slides.Count }
        }
        "export_video" {
          $outputPath = [string]$operation.outputPath
          $useTimings = [bool](Get-Property $operation "useTimingsAndNarrations" $true)
          $defaultDuration = [int](Get-Property $operation "defaultSlideDurationSeconds" 5)
          $resolution = [int](Get-Property $operation "verticalResolution" 1080)
          $framesPerSecond = [int](Get-Property $operation "framesPerSecond" 30)
          $quality = [int](Get-Property $operation "quality" 85)
          $timeoutSeconds = [int](Get-Property $operation "timeoutSeconds" 600)
          $presentation.CreateVideo($outputPath, $useTimings, $defaultDuration, $resolution, $framesPerSecond, $quality)
          $startedAt = Get-Date
          do {
            Start-Sleep -Milliseconds 250
            $status = [int]$presentation.CreateVideoStatus
            if ($status -eq 4) { throw "PowerPoint video export failed: $outputPath" }
            if (((Get-Date) - $startedAt).TotalSeconds -gt $timeoutSeconds) { throw "PowerPoint video export timed out after $timeoutSeconds seconds: $outputPath" }
          } while ($status -ne 3)
          if (-not (Test-Path -LiteralPath $outputPath)) { throw "PowerPoint reported a completed video export but no file was created: $outputPath" }
          $results += @{ op = $op; path = $outputPath; status = $status; resolution = $resolution; framesPerSecond = $framesPerSecond; quality = $quality }
        }
        default { throw "Unsupported PowerPoint operation: $op" }
      }
    }
    return @{ application = "powerpoint"; results = $results }
  } finally {
    if ($session.openedHere) { try { $presentation.Close() } catch {} }
    if ($session.ownedApp) {
      try {
        if ([int]$app.Windows.Count -eq 0 -and [int]$app.Presentations.Count -eq 0) { $app.Quit() }
      } catch {}
    }
    Release-ComObject $presentation; Release-ComObject $app
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
  }
}

try {
  $request = Read-JsonInput
  $result = switch ($Command) {
    "status" { Get-OfficeStatus }
    "list-open" { Get-OpenOfficeFiles }
    "excel-batch" { Invoke-ExcelBatch $request }
    "word-batch" { Invoke-WordBatch $request }
    "powerpoint-batch" { Invoke-PowerPointBatch $request }
  }
  Write-JsonOutput @{ ok = $true; result = $result }
  exit 0
} catch {
  Write-JsonOutput @{ ok = $false; error = @{ message = $_.Exception.Message; type = $_.Exception.GetType().FullName; command = $Command; line = $_.InvocationInfo.ScriptLineNumber; statement = $_.InvocationInfo.Line; stack = $_.ScriptStackTrace } }
  exit 1
}
