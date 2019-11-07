#zlicza udane testy z danego tygodznia


#dane do edycji:
#czy wyœwietlaæ dodatkowe informacje, bêdzie dzia³aæ wolniej
$debug = 0

#przciski, przesuniêcie rzêdu
$Right_Row_Button = 800
$wielkosc_czcionki_okna = 10
$rozmiar_kolumn = 105
$wysokosc_okna = 500




#danych poni¿ej nie edytowaæ

#czas przeznaczony na pisanie
# 2019.08.19 - 7h
# 2019.08.20 - 4h
# 2019.08.20 - 1,5h - zmiana na lepsz¹ tabele
# 2019.10.24 - 2h
# 2019.10.28 - 1,5h - dodanie sortowanie po kolumnach
# 2019.11.07 - 4h - FTT, export i import

$title = "Testy na Pass GUI wersja. 7E"

#przechowuje dane pobrane z plików
$Wynik = [ordered]@{}

$regPath="HKCU:\SOFTWARE\Rysiu22\TnP7C"
$name="path"
$regYear="rok"
$regPastWeek="od_tygodnia"
$regToWeek="do_tygodnia"

#folder z logami
$sciezka=[System.IO.Path]::GetDirectoryName($script:MyInvocation.MyCommand.Path) #aktualna œcie¿ka
$testRok="2019"
$od_t="1"
$do_t="52"

#wczytanie okienek
[System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") | Out-Null

IF( (Test-Path $regPath))
{
    #do poprawy
    $sciezka=(Get-Item -Path $regPath).GetValue($name)
	$testRok=(Get-Item -Path $regPath).GetValue($regYear)
	$od_t=(Get-Item -Path $regPath).GetValue($regPastWeek)
	$do_t=(Get-Item -Path $regPath).GetValue($regToWeek)
}
ELSE
{
	[System.Windows.Forms.MessageBox]::Show("Pierwsze uruchomienie! Ustaw poprawnie wszystkie pola i folder. Nastêpnie wciœnij Generuj",'Info')
}

#POBIERA AKTUALN¥ DATE
$dzien=get-date -UFormat "%Y-%m-%d"

#Tworzenie okna programu
Add-Type -AssemblyName System.Windows.Forms
$form = New-Object System.Windows.Forms.Form
$form.Text=$title
$form.Size=New-Object System.Drawing.Size(($Right_Row_Button+200), ($wysokosc_okna+120))
$form.StartPosition='CenterScreen'
#$form.topmost = $true

#wczytuje rozmiar czcionek
$MyFont = New-Object System.Drawing.Font("Lucida Console",$wielkosc_czcionki_okna,[System.Drawing.FontStyle]::Regular)
#$MyFont = New-Object System.Drawing.Font("Courier New",$wielkosc_czcionki_okna,[System.Drawing.FontStyle]::Regular)

#MENU BAR
#https://social.technet.microsoft.com/Forums/en-US/52debd7a-1f2b-470e-9259-c898563bb3ae/tool-strip-menu-item?forum=ITCG
$MenuBar = New-Object System.Windows.Forms.MenuStrip
$Form.Controls.Add($MenuBar)
$UserGMenu1 = New-Object System.Windows.Forms.ToolStripMenuItem
$UserGMenu2 = New-Object System.Windows.Forms.ToolStripMenuItem
$MenuBar.Items.Add($UserGMenu1)
$MenuBar.Items.Add($UserGMenu2)
$UserGMenu1.Text = "&Plik"
$UserGMenu2.Text = "&Akcja"
$UserGMenu1.Font = $MyFont
$UserGMenu2.Font = $MyFont

$DropDownGUsers1Dict=@{'Exportuj do JSON'={dojson}; 'Importuj z JSON'={zjson}}
ForEach ($GroupUserKey in ($DropDownGUsers1Dict.keys | Sort-Object)) {
	#Write-Host $GroupUserKey, $DropDownGUsers1Dict[$GroupUserKey]
	$GroupValue = New-Object System.Windows.Forms.ToolStripMenuItem
	$GroupValue.Text = $GroupUserKey
	# name the control
	$Groupvalue.Name = $GroupUserKey
	$UserGMenu1.DropDownItems.Add($GroupValue)
	# use name to identify control
	$GroupValue.Add_Click( $DropDownGUsers1Dict[$GroupUserKey] )
}

$DropDownGUsers2Dict=@{'Odœwie¿'={Odswiez}; 'Generuj'={Dzialaj}; 'Zmieñ Folder'={ChangeFolder} }

ForEach ($GroupUserKey in ($DropDownGUsers2Dict.keys | Sort-Object)) {
	#Write-Host $GroupUserKey, $DropDownGUsers2Dict[$GroupUserKey]
	$GroupValue = New-Object System.Windows.Forms.ToolStripMenuItem
	$GroupValue.Text = $GroupUserKey
	# name the control
	$Groupvalue.Name = $GroupUserKey
	$UserGMenu2.DropDownItems.Add($GroupValue)
	# use name to identify control
	$GroupValue.Add_Click( $DropDownGUsers2Dict[$GroupUserKey] )
}


#1 linia
$label1=New-Object System.Windows.Forms.label
$label1.Text="..."
$label1.AutoSize=$True
$label1.Top="30"
$label1.Left="10"
$label1.Anchor="Left,Top"
$label1.Font = $MyFont
$form.Controls.Add($label1)

#2 linia
$label2=New-Object System.Windows.Forms.label
$label2.Text="Rok"
$label2.AutoSize=$True
$label2.Top="55"
$label2.Left=($Right_Row_Button+10)
$label2.Anchor="Left,Top"
$label2.Font = $MyFont
$form.Controls.Add($label2)

#4 linia
$label2=New-Object System.Windows.Forms.label
$label2.Text="Od tygodnia"
$label2.AutoSize=$True
$label2.Top="105"
$label2.Left=$Right_Row_Button
$label2.Anchor="Left,Top"
$label2.Font = $MyFont
$form.Controls.Add($label2)

#5 linia
$label2=New-Object System.Windows.Forms.label
$label2.Text="Do tygodnia"
$label2.AutoSize=$True
$label2.Top="155"
$label2.Left=$Right_Row_Button
$label2.Anchor="Left,Top"
$label2.Font = $MyFont
$form.Controls.Add($label2)

#6 linia
$label6=New-Object System.Windows.Forms.label
$label6.Text="£¹cznie wyników: 0"
$label6.AutoSize=$True
$label6.Top="205"
$label6.Left=$Right_Row_Button
$label6.Anchor="Left,Top"
$label6.Font = $MyFont
$form.Controls.Add($label6)

#7 linia
$label7=New-Object System.Windows.Forms.label
$label7.Text="Aktualny tydzieñ: " + (get-date -UFormat %V)
$label7.AutoSize=$True
$label7.Top="225"
$label7.Left=$Right_Row_Button
$label7.Anchor="Left,Top"
$label7.Font = $MyFont
$form.Controls.Add($label7)


#OKNO 1
$listBox=New-Object System.Windows.Forms.Listbox
$listBox.Location = New-Object System.Drawing.Size(10,55)
$listBox.Size= New-Object System.Drawing.Size(($Right_Row_Button - 20),100)
$listbox.HorizontalScrollbar = $true;
$listBox.Font = $MyFont
#$form.Controls.Add($listBox)




#OKNO Z KOLUMNAMI
$listView = New-Object System.Windows.Forms.ListView
$ListView.Location = New-Object System.Drawing.Point(10, 55)
$ListView.Size = New-Object System.Drawing.Size(($Right_Row_Button - 20),$wysokosc_okna)
$ListView.View = [System.Windows.Forms.View]::Details
$ListView.FullRowSelect = $true;


#sortowanie kolumn

#https://stackoverflow.com/questions/35871501/listview-sort-doesnt-work-onclick-powershell
#https://www.soinside.com/question/qjdFNSnTeRzPA65VFQUHwN 
$tmp = "
function SortListView {
    Param(
        [System.Windows.Forms.ListView]$sender,
        $column
    )
    $temp = $sender.Items | Foreach-Object { $_ }
    $Script:SortingDescending = !$Script:SortingDescending
    $sender.Items.Clear()
    $sender.ShowGroups = $false
    $sender.Sorting = 'none'
    $sender.Items.AddRange(($temp | Sort-Object -Descending:$script:SortingDescending -Property @{ Expression={ $_.SubItems[$column].Text } }))
}
$ListView.add_ColumnClick({SortListView $this $_.Column})
"

#https://social.technet.microsoft.com/Forums/scriptcenter/de-DE/553f06bc-522c-4854-9e28-d0e219a789a6/powershell-and-systemwindowsformslistview?forum=ITCG

# This is the custom comparer class string
# copied from the MSDN article

$comparerClassString = @"

  using System;
  using System.Windows.Forms;
  using System.Drawing;
  using System.Collections;

  public class ListViewItemComparer : IComparer
  {
    private int col;
    public ListViewItemComparer()
    {
      col = 0;
    }
    public ListViewItemComparer(int column)
    {
      col = column;
    }
    public int Compare(object x, object y)
    {
		int number_1, number_2;
		if(Int32.TryParse(((ListViewItem)x).SubItems[col].Text,out number_1) && Int32.TryParse(((ListViewItem)y).SubItems[col].Text,out number_2))
		{
			return number_1.CompareTo(number_2);
		}
		else
		{
			return String.Compare(((ListViewItem)x).SubItems[col].Text, ((ListViewItem)y).SubItems[col].Text);
		}
    }
  }

"@

# Add the comparer class

Add-Type -TypeDefinition $comparerClassString -ReferencedAssemblies ('System.Windows.Forms', 'System.Drawing')

# Add the event to the ListView ColumnClick event
$ListView.add_ColumnClick({ $listView.ListViewItemSorter = New-Object ListViewItemComparer($_.Column)})




$ListView.Font = $MyFont
$form.Controls.Add($ListView)

$MyTextAlign = [System.Windows.Forms.HorizontalAlignment]::Right;

#Nazwy kolumn
$LVcol1 = New-Object System.Windows.Forms.ColumnHeader
$LVcol1.TextAlign = $MyTextAlign
$LVcol1.Text = "Folder"
$LVcol1.Width = $rozmiar_kolumn

$LVcol2 = New-Object System.Windows.Forms.ColumnHeader
$LVcol2.TextAlign = $MyTextAlign
$LVcol2.Text = "Tydzieñ"

$LVcol3 = New-Object System.Windows.Forms.ColumnHeader
$LVcol3.TextAlign = $MyTextAlign
$LVcol3.Text = "Rok"

$LVcol4 = New-Object System.Windows.Forms.ColumnHeader
$LVcol4.TextAlign = $MyTextAlign
$LVcol4.Text = "FPY - first pass yield"

$LVcol5 = New-Object System.Windows.Forms.ColumnHeader
$LVcol5.TextAlign = $MyTextAlign
$LVcol5.Text = "FTT - first total test"

$LVcol6 = New-Object System.Windows.Forms.ColumnHeader
$LVcol6.TextAlign = $MyTextAlign
$LVcol6.Text = "PY - pass yield"

$LVcol7 = New-Object System.Windows.Forms.ColumnHeader
$LVcol7.TextAlign = $MyTextAlign
$LVcol7.Text = "Modu³ów Suma"
$LVcol7.Width = $rozmiar_kolumn

$LVcol8 = New-Object System.Windows.Forms.ColumnHeader
$LVcol8.TextAlign = $MyTextAlign
$LVcol8.Text = "Pass Suma"
$LVcol8.Width = $rozmiar_kolumn

$LVcol9 = New-Object System.Windows.Forms.ColumnHeader
$LVcol9.TextAlign = $MyTextAlign
$LVcol9.Text = "Testów Suma"
$LVcol9.Width = $rozmiar_kolumn


$ListView.Columns.AddRange([System.Windows.Forms.ColumnHeader[]](@($LVcol1, $LVcol2, $LVcol3, $LVcol4, $LVcol5, $LVcol6, $LVcol7, $LVcol8, $LVcol9)))

#dzia³a dobrze
#$ListViewItem = New-Object System.Windows.Forms.ListViewItem([System.String[]](@("ISA", "52", "2019", "0","1", "6", "7", "8")), -1)
#$ListViewItem.StateImageIndex = 0
#$ListView.Items.AddRange([System.Windows.Forms.ListViewItem[]](@($ListViewItem)))	

#slabo dzia³a
#$import = @("ISA", "52", "2019", "0","1", "6", "7", "8")
#ForEach($array in $import){	
#	$item = New-Object System.Windows.Forms.ListviewItem($array)
#	$listView.Items.Add($item)}


#BUTTON
#GENERUJ
#$generate=New-Object System.Windows.Forms.Button
#$generate.Location=New-Object System.Drawing.Size(($Right_Row_Button+10),55)
#$generate.Size=New-Object System.Drawing.Size(100,30)
#$generate.Text="Generuj"
#$generate.add_click({Dzialaj})
#$form.Controls.Add($generate)


#CHECKBOX 1
$checkMe1=New-Object System.Windows.Forms.CheckBox
$checkMe1.Location=New-Object System.Drawing.Size(($Right_Row_Button+210),25)
$checkMe1.Size=New-Object System.Drawing.Size(100,30)
$checkMe1.Text="Debug"
$checkMe1.TabIndex=1
$checkMe1.Checked=$false
$checkMe1.Font = $MyFont
$form.Controls.Add($checkMe1)

#CHECKBOX 2
$checkMe2=New-Object System.Windows.Forms.CheckBox
$checkMe2.Location=New-Object System.Drawing.Size(($Right_Row_Button+10),425)
$checkMe2.Size=New-Object System.Drawing.Size(100,30)
$checkMe2.Text="Dopisuj wyniki"
$checkMe2.TabIndex=1
$checkMe2.Checked=$false
$checkMe2.Font = $MyFont
#$form.Controls.Add($checkMe2)

$textBoxPadingRight = 110

#TEXTBOX 1
$textBox1 = New-Object System.Windows.Forms.TextBox
$textBox1.Location = New-Object System.Drawing.Point(($Right_Row_Button+$textBoxPadingRight),55)
$textBox1.Size = New-Object System.Drawing.Size(40,30)
$textBox1.Text=$testRok
$textBox1.Font = $MyFont
$form.Controls.Add($textBox1)

#TEXTBOX 2
$textBox2 = New-Object System.Windows.Forms.TextBox
$textBox2.Location = New-Object System.Drawing.Point(($Right_Row_Button+$textBoxPadingRight),105)
$textBox2.Size = New-Object System.Drawing.Size(40,30)
$textBox2.Text=$od_t
$textBox2.Font = $MyFont
$form.Controls.Add($textBox2)

#TEXTBOX 3
$textBox3 = New-Object System.Windows.Forms.TextBox
$textBox3.Location = New-Object System.Drawing.Point(($Right_Row_Button+$textBoxPadingRight),155)
$textBox3.Size = New-Object System.Drawing.Size(40,30)
$textBox3.Text=$do_t
$textBox3.Font = $MyFont
$form.Controls.Add($textBox3)


#IMAGES
[System.Windows.Forms.Application]::EnableVisualStyles();


#zwraca iloœc modu³ów /del
function Zliczaj1($co)
{
	$B=@()
    foreach($i in $co)
    {
		$D = ($I.NAME -replace '(.+)_\d+.txt','$1')
        $B += $D
    }
    $C = ($B | Group-Object -NoElement)
    return $c
}


#zwraca liste plików z ostatnimi testami modu³ów
function Zliczaj2($co)
{
	#tworzy slownik[sn_modulu] = (max_tygodniowy_numer_testu, dane_do_pliku)
	# iloœc modu³ów to: slownik.keys.COUNT
	$B=@{}
    foreach($i in $co)
    {
		$D = ($I.NAME -replace '(.+)_\d+.txt','$1')
		#$B += $D
		
		$D_last = ($I.NAME -replace '.+_(\d+).txt','$1')
		#write-host $D, $D_last
		
		if($D_last -match '^\d+$')
		{
			if($B[$D].Length -eq 0)
			{
				$B[$D] = @($D_last,$I)
			}
			else
			{
				if([convert]::ToInt32($B[$D][0],10) -lt [convert]::ToInt32($D_last,10))
				{
					$B[$D] = @($D_last,$I)
				}
			}		
		}
		else
		{
			write-host "to nie jest liczba na koncu pliku: '$D_last'" 
		}
    }

	$C = @()
	foreach($i in $B.keys)
	{
		if($debug){write-host "{$i : ",$B[$i][0],$B[$i][1].Fullname}
		$C += $B[$i][1]
	}
	
	#zwraca liste: dane_do_pliku z ostatnim testem
	return $C
}




#wczytanie i tworzenie danych
function GetList($sciezka1)
{

	$testRok=$textBox1.Text
	$od_t=$textBox2.Text
	$do_t=$textBox3.Text
	
	#zmienna z danymi
	$Dict = @{}
	$Result = @{}

	#pobranie posortowanych plików
	$pliki = (Get-ChildItem $sciezka1 *.TXT) # | sort LastWriteTime

	#wype³nianie zmiennej danymi na temat logów {rok:{tydzien: [dane] }
	foreach($plik in $pliki)
	{
		$rok = (get-date $plik.LastWriteTime -UFormat %Y)
		$plik_tydzien = (get-date $plik.LastWriteTime -UFormat %V)
		#if($debug){write-host $plik, (get-date $plik.LastWriteTime -UFormat "%Y.%V")}
		
		#jeœli rok = 0 to pomija wszelkie restrykcjie czasowe
		if($testRok -ne "0")
		{
			if($rok -ne $testRok)
			{
				continue
			}
			
			if([convert]::ToInt32($plik_tydzien,10) -lt $od_t)
			{
				#write-host [convert]::ToInt32($plik_tydzien,10) , "-lt", $od_t
				continue
			}
			if([convert]::ToInt32($plik_tydzien,10) -gt $do_t)
			{
				#write-host [convert]::ToInt32($plik_tydzien,10), "-gt", $do_t
				continue
			}
		}
		
		if($Dict[$rok].Length -eq 0)
		{
			$Dict[$rok] = @{}
		}

		if($Dict[$rok][$plik_tydzien].Length -eq 0)
		{
			$Dict[$rok][$plik_tydzien] = @($plik)
		}
		else
		{
			$Dict[$rok][$plik_tydzien] += $plik
		}
		#write-host "check_end",$Dict[$plik_tydzien].Length
	}

	#wyœwietlenie struktury rok i tydzieñ
	if($debug)
	{
		foreach($year in ($Dict.keys)) # | Sort-Object {[double]$_}))
		{
			write-host "debug key{$year : ... } count value:", $Dict[$year].Length
			
			foreach($week in ($Dict[$year].keys)) #  | Sort-Object {[double]$_}))
			{
				write-host "debug key{$year : {$week : ... }} count value:", $Dict[$year].$week.Length
				
				foreach($i in ($Dict[$year][$week]))
				{
					write-host "debug value: $i"
				}
			}
		}
		write-host "###############################"
		write-host ""
	}

	#Przetworzenie zebranych danych
	foreach($year in ($Dict.keys)) # | Sort-Object {[double]$_}))
	{	
		if($checkMe1.Checked){write-host "key{$year : ... } count value:", $Dict[$year].Length}
		
		foreach($week in ($Dict[$year].keys)) # | Sort-Object {[double]$_}))
		{
			#write-host "key{$year : {$week : ... }} count value:", $Dict[$year].$week.Length
			#write-host $Dict[$year].$week

			#poprzednia wersja sprawdzania testu na pass, PY
			#$lista=@($Dict[$year].$week | WHERE-OBJECT { $A=GET-CONTENT $_.FULLNAME; $A -MATCH "FAILS=0" }  | WHERE-OBJECT { $A=GET-CONTENT $_.FULLNAME; $A -MATCH "ERRORS=0" } )
			#write-host "rok:$year tydzien:$week PY:", $lista.Length, "/", $Dict[$year].$week.Length

			#FTT
			$lista_first=@($Dict[$year].$week | WHERE-OBJECT { $A=GET-CONTENT $_.FULLNAME -Head 10; $A -MATCH "RESULT="} | WHERE-OBJECT { $_.Name -MATCH "_0.txt" } )

			#FPY
			$lista_first_pass=@($Dict[$year].$week | WHERE-OBJECT { $A=GET-CONTENT $_.FULLNAME -Head 10; $A -MATCH "RESULT=PASS"} | WHERE-OBJECT { $_.Name -MATCH "_0.txt" } )
			
			#PY
			#sprawdzenie czy dany log zawiera ci¹g znaków w pierwszych 10 liniach, jesli tak to test zaliczany jako pass
			#dodaæ sprawdzanie czy plik zawieraj¹ obie linie !
			$lista_pass=@($Dict[$year].$week | WHERE-OBJECT { $A=GET-CONTENT $_.FULLNAME -Head 10; $A -MATCH "RESULT=PASS" } )
			$lista_fail=@($Dict[$year].$week | WHERE-OBJECT { $A=GET-CONTENT $_.FULLNAME -Head 10; $A -MATCH "RESULT=FAIL" } )
			$znalezione_testy = $lista_pass.Length + $lista_fail.Length
			if($checkMe1.Checked){write-host "rok:$year tydzien:$week FPY:", $lista_first_pass.Length, "/", $znalezione_testy, "PY:", $lista_pass.Length, "/", $znalezione_testy}
			
			
			$lista_last_test=(Zliczaj2 $Dict[$year][$week])
			#write-host $lista_last_test
			$lista_last_pass=@($lista_last_test | WHERE-OBJECT { $A=GET-CONTENT $_.FULLNAME -Head 10; $A -MATCH "RESULT=PASS" } )
			#write-host "lista_last_pass",$lista_last_pass.Length
			
			#ile modu³ów zosta³o przetestowanych
			$ile_mod=$lista_last_test.COUNT;
			
			#wykrycie b³êdu w obliczeniach
			if(($lista_pass.Length + $lista_fail.Length) -ne $Dict[$year].$week.Length)
			{
				write-host "Uwaga!!! Wykryto plik ktory nie jest logiem z testow! tydzien:$week PASS + FAIL =", $lista_pass.Length, "+", $lista_fail.Length,"=",$Dict[$year].$week.Length
				write-host ""
				
			}
			
			if($Result[$year].Length -eq 0)
			{
				$Result[$year] = @{}
			}

			if($Result[$year][$week].Length -eq 0)
			{
				$Result[$year][$week] = @{"FPY"=$lista_first_pass.Length; "PY"=$lista_last_pass.Length ;"sum_pass"= $lista_pass.Length; "sum_test"= $znalezione_testy; "sum_moduly"= $ile_mod; "pliki_pass"= $lista_pass | Select-Object -Property Name; "pliki_fail"= $lista_fail | Select-Object -Property Name; "pliki_first_pass"=$lista_first_pass | Select-Object -Property Name; "pliki_last_pass"= $lista_last_pass | Select-Object -Property Name; "FTT"= $lista_first.Length}
				#write-host $Result[$year][$week]["tydzien"]
			}
			else
			{
				write-host "Uwaga!!! Wykryto b³¹d spójnoœci testów"
			}
			
		}
		if($checkMe1.Checked){write-host "koniec $year"}
		#write-host "week",$Result[$year].keys
	}
	if($checkMe1.Checked){write-host "koniec $sciezka1"}
	#write-host $Result.keys

	return ,$Result
}


#https://stackoverflow.com/questions/40495248/create-hashtable-from-json
#rekurencyjny poprawny import zmiennych
#[CmdletBinding]
function Get-FromJson
{
    param(
        [Parameter(Mandatory=$true, Position=1)]
        [string]$Path
    )

    function Get-Value {
        param( $value )

        $result = $null
        if ( $value -is [System.Management.Automation.PSCustomObject] )
        {
            Write-Verbose "Get-Value: value is PSCustomObject"
            $result = @{}
            $value.psobject.properties | ForEach-Object { 
                $result[$_.Name] = Get-Value -value $_.Value
				#write-host "-" $_.Name
            }
        }
        elseif ($value -is [System.Object[]])
        {
            $list = New-Object System.Collections.ArrayList
            Write-Verbose "Get-Value: value is Array"
            $value | ForEach-Object {
                $list.Add((Get-Value -value $_)) | Out-Null
            }
            $result = $list
        }
        else
        {
            Write-Verbose "Get-Value: value is type: $($value.GetType())"
            $result = $value
        }
        return $result
    }


    if (Test-Path $Path)
    {
        $json = Get-Content $Path -Raw
    }
    else
    {
        $json = '{}'
    }

    $hashtable = Get-Value -value (ConvertFrom-Json $json)

    return $hashtable
}

#https://gallery.technet.microsoft.com/scriptcenter/GUI-popup-FileSaveDialog-813a4966
function dojson()
{
    $openDiag=New-Object System.Windows.Forms.savefiledialog
	$openDiag.initialDirectory = [System.IO.Directory]::GetCurrentDirectory()
	$openDiag.filter = "Log Files|*.json|All Files|*.*" 
    #otwieranie ostatnio wybrany folder
    $result=$openDiag.ShowDialog()
    if($result -eq "OK")
    {
	    #[System.Windows.Forms.MessageBox]::Show("Ustawiono: "+$openDiag.filename,'plik')
		#$Wynik | Select-Object -Property * | ConvertTo-JSON -Depth 4 | Set-Content -Path $openDiag.filename
		$Wynik | ConvertTo-JSON -Depth 4 | Set-Content -Path $openDiag.filename
		#$Wynik | ForEach-OBJECT{ [pscustomobject]$_} | Export-CSV -Path "dump.csv"
    }
}

#https://community.idera.com/database-tools/powershell/powertips/b/tips/posts/using-open-file-dialogs
function zjson()
{
    $openDiag=New-Object System.Windows.Forms.OpenFileDialog
	$openDiag.Multiselect = $false
	$openDiag.initialDirectory = [System.IO.Directory]::GetCurrentDirectory()
	$openDiag.filter = "Log Files|*.json|All Files|*.*" 
    #otwieranie ostatnio wybrany folder
    $result=$openDiag.ShowDialog()
    if($result -eq "OK")
    {
	    #[System.Windows.Forms.MessageBox]::Show("Ustawiono: "+$openDiag.filename,'plik')
		#$script:Wynik = (Get-Content -Raw -Path $openDiag.filename | ConvertFrom-Json)
		#write-host ($Wynik | ConvertTo-JSON -Depth 4)
		
		$script:Wynik = Get-FromJson $openDiag.filename
		
		Odswiez
	}
}


#akcje Generowania
function Dzialaj()
{
	write-host "Start"
	zapis_konfiguracji
	
	#zerowanie zmiennej
	$script:Wynik = [ordered]@{}
	
	foreach($path in (Get-ChildItem $sciezka))
	{
		write-host $path,$path.Name
		if((Get-Item $path.FULLNAME) -is [System.IO.DirectoryInfo])
		{
			if($checkMe1.Checked){write-host $path.FULLNAME}
			$Wynik[$path.Name] = GetList($path.FULLNAME)
		}
	}
	
	Odswiez

	$date = get-date
	write-host "Koniec", $date
}

function zapis_konfiguracji()
{
	#zapis do rejstru
    IF( ! (Test-Path $regPath))
    {
		New-Item -Path $regPath -Force | Out-Null
	}	
	$testRok=$textBox1.Text
	$od_t=$textBox2.Text
	$do_t=$textBox3.Text
	
	if($debug){write-host "$od_t-$do_t/$testRok"}
    
	New-ItemProperty -Path $regPath -Name $regYear -Value $testRok -Force | Out-Null
	New-ItemProperty -Path $regPath -Name $regPastWeek -Value $od_t -Force | Out-Null
	New-ItemProperty -Path $regPath -Name $regToWeek -Value $do_t -Force | Out-Null
	New-ItemProperty -Path $regPath -Name $name -Value $sciezka -Force | Out-Null
}

function Odswiez()
{
	zapis_konfiguracji
	
	#odczytanie zmiennych w oknach
	$testRok=$textBox1.Text
	$od_t=$textBox2.Text
	$do_t=$textBox3.Text

	#odœwierzenie listy
	$listBox.Items.Clear()
	$listView.Items.Clear()
	
	$label6.Text="£¹cznie wyników: " + ($Wynik.keys).COUNT
	$label6.Refresh()

	foreach($modul in ($Wynik.keys))
	{
		if($checkMe1.Checked){write-host "modul", $modul}
		
		foreach($year in ($Wynik[$modul].keys | Sort-Object {[double]$_}))
		{
			if($checkMe1.Checked){write-host "key{$modul : {$year : ... }} count value:", $Wynik[$modul][$year].Length}

			foreach($week in ($Wynik[$modul][$year].keys | Sort-Object {[double]$_}))
			{
				if($checkMe1.Checked){write-host "key{$modul : {$year : {$week : ... }}} count value:", $Wynik[$modul][$year].$week.Length, $Wynik.$modul.$year.$week["FPY"],$Wynik.$modul.$year.$week["PY"],$Wynik.$modul.$year.$week["sum"]}
				
				#jeœli rok = 0 to pomija wszelkie restrykcjie czasowe
				if($testRok -ne "0")
				{
					#restrykcje czasowe: rok
					if($year -ne $testRok)
					{
						continue
					}
			
					#restrykcje czasowe: tygodnie
					if([convert]::ToInt32($week,10) -lt $od_t)
					{
						#write-host [convert]::ToInt32($plik_tydzien,10) , "-lt", $od_t
						continue
					}
					if([convert]::ToInt32($week,10) -gt $do_t)
					{
						#write-host [convert]::ToInt32($plik_tydzien,10), "-gt", $do_t
						continue
					}
				}
				
				#latwe do odczytania w kodzie
				#$listBox.Items.Add("$week/$year  $modul   FPY: "+ $Wynik.$modul.$year.$week["FPY"] + "   PY: "+ $Wynik.$modul.$year.$week["PY"] + " Suma pass: " + $Wynik.$modul.$year.$week["sum_pass"] + " Suma mod: " + $Wynik.$modul.$year.$week["sum_moduly"] + " Suma_testow: " + $Wynik.$modul.$year.$week["sum_test"] )
				
				#latwe do odczytania w oknie
				#$pad = 3
				#$week_f = ([convert]::ToInt32($week,10).ToString("#00")).PadLeft(3)
				#$listBox.Items.Add("$modul $week_f/$year FPY: "+ ($Wynik.$modul.$year.$week["FPY"].ToString("#0")).PadLeft($pad) + "   PY: "+ ($Wynik.$modul.$year.$week["PY"].ToString("#0")).PadLeft($pad) + " Suma mod: " + ($Wynik.$modul.$year.$week["sum_moduly"].ToString("#0")).PadLeft($pad) + " Suma pass: " + ($Wynik.$modul.$year.$week["sum_pass"].ToString("#0")).PadLeft($pad) + " Suma_testow: " + ($Wynik.$modul.$year.$week["sum_test"].ToString("#0")).PadLeft($pad) )

				#wype³nanie tabeli
				$ListViewItem = New-Object System.Windows.Forms.ListViewItem([System.String[]](@($modul, $week, $year, $Wynik.$modul.$year.$week["FPY"], $Wynik.$modul.$year.$week["FTT"], $Wynik.$modul.$year.$week["PY"], $Wynik.$modul.$year.$week["sum_moduly"], $Wynik.$modul.$year.$week["sum_pass"], $Wynik.$modul.$year.$week["sum_test"])), -1)
				#$ListViewItem.StateImageIndex = 0
				$ListView.Items.AddRange([System.Windows.Forms.ListViewItem[]](@($ListViewItem)))	
			}
		}
		
	}
	
	#$label6.Text="Wyników: " + ($ListView.Items).COUNT
	#$label6.Refresh()
	
	if(($Wynik.keys).COUNT -eq 0)
	{
		UpdateList
		
	}
}

function UpdateList()
{
    #okno
    $listBox.Items.Clear()
	$listView.Items.Clear()
	$label1.Text=$sciezka
	$label1.Refresh()
	
	foreach($path in (Get-ChildItem $sciezka))
	{
		if((Get-Item $path.FULLNAME) -is [System.IO.DirectoryInfo])
		{
			$listBox.Items.Add($path.NAME) 
		}
	}
}

function ChangeFolder()
{
    $openDiag=New-Object System.Windows.Forms.folderbrowserdialog
    #$openDiag.rootfolder="MyComputer"
    $openDiag.rootfolder="Desktop"
    #otwieranie ostatnio wybrany folder
    $openDiag.Description="Wybierz folder z logami"
    $openDiag.SelectedPath = $sciezka
    $result=$openDiag.ShowDialog()
    if($result -eq "OK")
    {
        $script:sciezka = $openDiag.SelectedPath
	    #[System.Windows.Forms.MessageBox]::Show("Ustawiono: "+$sciezka,'folder')
        $label1.Text=$sciezka

        #UpdateList
		$label1.Text=$sciezka
		$label1.Refresh()
    }
}

UpdateList

#pokazanie okna
$form.ShowDialog()
