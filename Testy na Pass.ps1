#zlicza udane testy z danego tygodznia


function Get-WeekNumber([datetime]$DateTime = (Get-Date)) 
{
	$ci = [System.Globalization.CultureInfo]::CurrentCulture
	($ci.Calendar.GetWeekOfYear($DateTime,$ci.DateTimeFormat.CalendarWeekRule,$ci.DateTimeFormat.FirstDayOfWeek)).ToString()
}


#dane do edycji:
#czy wyœwietlaæ dodatkowe informacje, bêdzie dzia³aæ wolniej
$debug = 0

#przciski, przesuniêcie rzêdu
$Right_Row_Button = 800
$wielkosc_czcionki_okna = 10
$rozmiar_kolumn = 105
$wysokosc_okna = 500

$plik_wzorcow = "wzorce_nazw_plikow.ini"

$wzorzec_karty = "\w+\d+-\d+_B\d+W\d+S\d+_\d+\.txt"
$wzorzec_wzmacniacza = "10P000ABT\d+_AMMAWZ\d{10}_\d+\.txt"

$ile_lini_czytac = 15


#danych poni¿ej nie edytowaæ

#czas przeznaczony na pisanie
# 2019.08.19 - 7h
# 2019.08.20 - 4h
# 2019.08.20 - 1,5h - zmiana na lepsz¹ tabele
# 2019.10.24 - 2h
# 2019.10.28 - 1,5h - dodanie sortowanie po kolumnach
# 2019.11.07 - 4h - FTT, export i import
# 2019.11.08 - 4h
# 2019.11.09 - 9,5h - wczytanie kompletnych danych z nag³óka i generowanie z nich danych, filtrowanie nazw tylko przy generowaniu
# 2019.11.11 - 9h
# 2019.12.03 - 3,5h - 19:00-22:30 dodano wczytywanie wzorców z osobnego pliku, poprawienie kolorów podczas sortowania, suma tygodni tylko podczas ³adowania, testy z klinanym menu
# 2020.02.16 - 2,5h

$title = "Testy na Pass GUI wersja. 7G"

#przechowuje dane pobrane z plików
$Wynik = [ordered]@{}

$regPath="HKCU:\SOFTWARE\Rysiu22\TnP7F"
$name="path"
$regYear="rok"
$regPastWeek="od_tygodnia"
$regToWeek="do_tygodnia"
$regMyRegxFile="filter_plikow"

#folder z logami
$sciezka=[System.IO.Path]::GetDirectoryName($script:MyInvocation.MyCommand.Path) #aktualna œcie¿ka
$testRok="2019"
$od_t="1"
$do_t="52"
$myRegxFile=".*"

#wczytanie okienek
[System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") | Out-Null

IF( (Test-Path $regPath))
{
    #do poprawy
    $sciezka=(Get-Item -Path $regPath).GetValue($name)
	$testRok=(Get-Item -Path $regPath).GetValue($regYear)
	$od_t=(Get-Item -Path $regPath).GetValue($regPastWeek)
	$do_t=(Get-Item -Path $regPath).GetValue($regToWeek)
	$myRegxFile=(Get-Item -Path $regPath).GetValue($regMyRegxFile)
	
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
$form.Size=New-Object System.Drawing.Size(($Right_Row_Button+300), ($wysokosc_okna+120))
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
$UserGMenu3 = New-Object System.Windows.Forms.ToolStripMenuItem
$UserGMenu4 = New-Object System.Windows.Forms.ToolStripMenuItem
$UserGMenu5 = New-Object System.Windows.Forms.ToolStripMenuItem
$MenuBar.Items.Add($UserGMenu1) | Out-Null
$MenuBar.Items.Add($UserGMenu2) | Out-Null
$MenuBar.Items.Add($UserGMenu3) | Out-Null
$MenuBar.Items.Add($UserGMenu4) | Out-Null
$MenuBar.Items.Add($UserGMenu5) | Out-Null
$UserGMenu1.Text = "&Plik"
$UserGMenu2.Text = "&Akcja"
$UserGMenu3.Text = "&Wzorzec Nazw"
$UserGMenu4.Text = "&Okno"
$UserGMenu5.Text = "&Info"
$UserGMenu1.Font = $MyFont
$UserGMenu2.Font = $MyFont
$UserGMenu3.Font = $MyFont
$UserGMenu4.Font = $MyFont
$UserGMenu5.Font = $MyFont

$DropDownGUsers1Dict=@{
	'2 Exportuj do GZip'={dojson -Format "gz"}; 
	'3 Exportuj do JSON'={dojson -Format "json"}; 
	'4 Exportuj do base64'={dojson -Format "base64"}; 
	
	'5 Importuj z GZip'={zjson -Format "gz"}; 
	'6 Importuj z JSON'={zjson -Format "json"}; 
	'7 Importuj z base64'={zjson -Format "base64"}; 

	"8 Zamknij"={$form.Close();}
}

ForEach ($GroupUserKey in ($DropDownGUsers1Dict.keys | Sort-Object)) {
	#Write-Host $GroupUserKey, $DropDownGUsers1Dict[$GroupUserKey]
	$GroupValue = New-Object System.Windows.Forms.ToolStripMenuItem
	$GroupValue.Text = $GroupUserKey.Substring(2)
	# name the control
	$Groupvalue.Name = $GroupUserKey
	$UserGMenu1.DropDownItems.Add($GroupValue) | Out-Null
	# use name to identify control
	$GroupValue.Add_Click( $DropDownGUsers1Dict[$GroupUserKey] )
}

$DropDownGUsers2Dict=@{
	#'1 Odœwie¿'={Odswiez}; 
	'2 £aduj'={Dzialaj}; 
	'3 Zmieñ Folder'={ChangeFolder}
	'4 Przelicz ponownie'={$Wynik = ObliczPonownie $Wynik; Odswiez};
	'5 Czyœæ Tabele'={$script:Wynik=@{}; Odswiez}
}

ForEach ($GroupUserKey in ($DropDownGUsers2Dict.keys | Sort-Object)) {
	#Write-Host $GroupUserKey, $DropDownGUsers2Dict[$GroupUserKey]
	$GroupValue = New-Object System.Windows.Forms.ToolStripMenuItem
	$GroupValue.Text = $GroupUserKey.Substring(2)
	# name the control
	$Groupvalue.Name = $GroupUserKey
	$UserGMenu2.DropDownItems.Add($GroupValue) | Out-Null
	# use name to identify control
	$GroupValue.Add_Click( $DropDownGUsers2Dict[$GroupUserKey] )
}

$DropDownGUsers3Dict=@{
	'1 Wszystko: .*' = {
		$script:myRegxFile=".*"; 
		$label8.Text="wzorzec: ",$myRegxFile; 
		$label8.Refresh();
		$Wynik = ObliczPonownie $Wynik; Odswiez;
		};
	"2 Karta: $wzorzec_karty" = {
		$script:myRegxFile=$wzorzec_karty;
		$label8.Text="wzorzec: ",$myRegxFile;
		$label8.Refresh();
		$Wynik = ObliczPonownie $Wynik; Odswiez;
		};
	"3 Wzmacniacz: $wzorzec_wzmacniacza" = {
		$script:myRegxFile=$wzorzec_wzmacniacza;
		$label8.Text="wzorzec: ",$myRegxFile; 
		$label8.Refresh();
		$Wynik = ObliczPonownie $Wynik; Odswiez;
		};
	'4 W³asny' = {
		$tmp=GetStringFromUser "Info" "Podaj w³asny wzorzec" $script:myRegxFile; 
		if($tmp){$script:myRegxFile=$tmp}; 
		$label8.Text="wzorzec: ",$myRegxFile; 
		$label8.Refresh();
		$Wynik = ObliczPonownie $Wynik; Odswiez;
		};
}

#czyta zawartoœæ pliku z wzorcami i dodaje wzorce do menu
If([System.IO.File]::Exists($plik_wzorcow))
{
'
	$wzorce = @(GET-CONTENT $plik_wzorcow | ForEach-Object{[Regex]::Escape($_) | Select-String -Pattern ".+=.*" } | ConvertFrom-StringData)
	$i = 5
	ForEach($key in $wzorce.keys)
	{
		#write-host ($key | Out-String)
		$tmpu = $key
		$key = ($i++).ToString() +" "+$key+": "+$wzorce.$key
		$DropDownGUsers3Dict.Add($key, {
			#write-host ($tmpu | Out-String);
			$script:myRegxFile=($wzorce.$key | Out-String);
			$label8.Text="wzorzec: ",($myRegxFile | Out-String);
			$label8.Refresh();
			$Wynik = ObliczPonownie $Wynik; Odswiez;
			}
			)
	}
'
}
ELSE
{
	write-host "Nie znaleziono pliku z wzorcami"
}

ForEach ($GroupUserKey in ($DropDownGUsers3Dict.keys | Sort-Object)) {
	#Write-Host $GroupUserKey, $DropDownGUsers2Dict[$GroupUserKey]
	$GroupValue = New-Object System.Windows.Forms.ToolStripMenuItem
	$GroupValue.Text = $GroupUserKey.Substring(2)
	# name the control
	$Groupvalue.Name = $GroupUserKey
	$UserGMenu3.DropDownItems.Add($GroupValue) | Out-Null
	# use name to identify control
	$GroupValue.Add_Click( $DropDownGUsers3Dict[$GroupUserKey] )
}

$DropDownGUsers4Dict=@{
	'1 Zawsze na wieszku'={$form.topmost = -not $form.topmost};
}

ForEach ($GroupUserKey in ($DropDownGUsers4Dict.keys | Sort-Object)) {
	#Write-Host $GroupUserKey, $DropDownGUsers1Dict[$GroupUserKey]
	$GroupValue = New-Object System.Windows.Forms.ToolStripMenuItem
	$GroupValue.Text = $GroupUserKey.Substring(2)
	# name the control
	$Groupvalue.Name = $GroupUserKey
	$UserGMenu4.DropDownItems.Add($GroupValue) | Out-Null
	# use name to identify control
	$GroupValue.Add_Click( $DropDownGUsers4Dict[$GroupUserKey] )
}

$DropDownGUsers5Dict=@{
	'1 Strona projektu'={Start-Process "https://github.com/Rysiu22/LicznikTestow"};
	'2 Pobierz Aktualn¹ wersje'={
		if( $script:MyInvocation.MyCommand.Path ){Invoke-WebRequest -Uri "https://raw.githubusercontent.com/Rysiu22/LicznikTestow/master/Testy%20na%20Pass.ps1" -OutFile $script:MyInvocation.MyCommand.Path};
		$form.Close();
		};
	
	
}

ForEach ($GroupUserKey in ($DropDownGUsers5Dict.keys | Sort-Object)) {
	#Write-Host $GroupUserKey, $DropDownGUsers1Dict[$GroupUserKey]
	$GroupValue = New-Object System.Windows.Forms.ToolStripMenuItem
	$GroupValue.Text = $GroupUserKey.Substring(2)
	# name the control
	$Groupvalue.Name = $GroupUserKey
	$UserGMenu5.DropDownItems.Add($GroupValue) | Out-Null
	# use name to identify control
	$GroupValue.Add_Click( $DropDownGUsers5Dict[$GroupUserKey] )
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
$label6.Text="£¹cznie folderów: 0"
$label6.AutoSize=$True
$label6.Top="205"
$label6.Left=$Right_Row_Button
$label6.Anchor="Left,Top"
$label6.Font = $MyFont
$form.Controls.Add($label6)

#7 linia
$label7=New-Object System.Windows.Forms.label
$label7.Text="Aktualny tydzieñ: " + (Get-WeekNumber) + " DATA: " + (get-date -UFormat "%Y-%m-%d")
$label7.AutoSize=$True
$label7.Top="225"
$label7.Left=$Right_Row_Button
$label7.Anchor="Left,Top"
$label7.Font = $MyFont
$form.Controls.Add($label7)

#7 linia
$label8=New-Object System.Windows.Forms.label
$label8.Text="wzorzec: ",$myRegxFile
$label8.AutoSize=$True
$label8.Top="255"
$label8.Left=$Right_Row_Button
$label8.Anchor="Left,Top"
$label8.Font = $MyFont
$form.Controls.Add($label8)

#OKNO Z KOLUMNAMI
$listView = New-Object System.Windows.Forms.ListView
$ListView.Location = New-Object System.Drawing.Point(10, 55)
$ListView.Size = New-Object System.Drawing.Size(($Right_Row_Button - 20),$wysokosc_okna)
$ListView.View = [System.Windows.Forms.View]::Details
$ListView.FullRowSelect = $true;



$contextMenuStrip1 = New-Object System.Windows.Forms.ContextMenuStrip
$contextMenuStrip1.Items.Add("Pliki").add_Click({Logi($ListView.SelectedItems.SubItems)})
$contextMenuStrip1.Items.Add("Kopiuj FP, FTT").add_Click(
{
	$item=$ListView.SelectedItems.SubItems;
	$tmp=$Wynik[$item[0].Text][$item[2].Text][$item[1].Text]
	if(($tmp["FPY"]).ToString() -ne $item[3].Text -or ($tmp["FTT"]).ToString() -ne $item[4].Text)
	{
		write-host("Uwaga skopiowana wartoœæ mo¿e byæ nie poprawna")
	}
	#(($tmp["FPY"]).ToString() + "	" + ($tmp["FTT"]).ToString() | Set-Clipboard)
	($item[3].Text + "	" + $item[4].Text | Set-Clipboard)
})

$contextMenuStrip1.Items.Add("Kopiuj PY, Modu³ów Suma").add_Click(
{
	$item=$ListView.SelectedItems.SubItems;
	$tmp=$Wynik[$item[0].Text][$item[2].Text][$item[1].Text]
	if(($tmp["PY"]).ToString() -ne $item[5].Text -or ($tmp["sum_moduly"]).ToString() -ne $item[6].Text)
	{
		write-host("Uwaga skopiowana wartoœæ mo¿e byæ nie poprawna")
	}
	#(($tmp["PY"]).ToString() + "	" + ($tmp["sum_moduly"]).ToString() | Set-Clipboard)
	($item[5].Text + "	" + $item[6].Text | Set-Clipboard)
})

$contextMenuStrip1.Items.Add("Kopiuj ca³y wiersz").add_Click(
{
	$item=$ListView.SelectedItems.SubItems;
	#write-host ($item | Select-Object -Property Text)
	($item | Select-Object -ExpandProperty Text) -join "`t" | Set-Clipboard
})

#$ListView.ContextMenu = $contextMenuStrip1
#$ListView.ShortcutsEnabled = $false
$ListView.ContextMenuStrip = $contextMenuStrip1

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
$ListView.add_ColumnClick({ $listView.ListViewItemSorter = New-Object ListViewItemComparer($_.Column); UstawoKolorWierszy($ListView) })


function UstawoKolorWierszy($ListView)
{
	For ($i=0; $i -lt $ListView.Items.Count; $i++)
	{
		if(($i % 2) -eq 0)
		{
			$ListView.Items[$i].BackColor = [System.Drawing.Color]::LightGray;
		}
		else
		{
			$ListView.Items[$i].BackColor = [System.Drawing.Color]::White;
		}
	}
}

#$ListView.Controls.Add($textBox1)

function Logi($item)
{
	if(! $item)
	{
		return
	}
	$fileContent = @{}

	$Wynik[$item[0].Text][$item[2].Text][$item[1].Text]["pliki"].GetEnumerator() | WHERE-OBJECT { $_.Name | Select-String -Pattern $myRegxFile } | ForEach-Object { $fileContent.Add($_.Name, $_.Value) }

	if($checkMe1.Checked){write-host ($fileContent | ConvertTo-JSON -Depth 2)}
	
	#Tworzenie okna programu
	Add-Type -AssemblyName System.Windows.Forms
	$form = New-Object System.Windows.Forms.Form
	$form.Text=$item[0].Text
	$form.Size=New-Object System.Drawing.Size(($Right_Row_Button+300), ($wysokosc_okna+120))
	$form.StartPosition='CenterScreen'

	#OKNO Z KOLUMNAMI
	$listView = New-Object System.Windows.Forms.ListView
	$ListView.Location = New-Object System.Drawing.Point(10, 55)
	$ListView.Size = New-Object System.Drawing.Size(($Right_Row_Button - 20 + 250),$wysokosc_okna)
	$ListView.View = [System.Windows.Forms.View]::Details
	$ListView.FullRowSelect = $true;
	$ListView.Font = $MyFont
	$form.Controls.Add($ListView)

	$MyTextAlign = [System.Windows.Forms.HorizontalAlignment]::Left;
	
	#"PN#0","SN#0","PN#1","SN#1","RESULT","START","USER","ERRORS","FAILS","SEQ_FILE","SEQ_MD5"

	#Nazwy kolumn
	$LVcol1 = New-Object System.Windows.Forms.ColumnHeader
	$LVcol1.TextAlign = $MyTextAlign
	$LVcol1.Text = "Nazwa"

	$LVcol2 = New-Object System.Windows.Forms.ColumnHeader
	$LVcol2.TextAlign = $MyTextAlign
	$LVcol2.Text = "PN#0"

	$LVcol3 = New-Object System.Windows.Forms.ColumnHeader
	$LVcol3.TextAlign = $MyTextAlign
	$LVcol3.Text = "SN#0"
	
	$LVcol4 = New-Object System.Windows.Forms.ColumnHeader
	$LVcol4.TextAlign = $MyTextAlign
	$LVcol4.Text = "PN#1"

	$LVcol5 = New-Object System.Windows.Forms.ColumnHeader
	$LVcol5.TextAlign = $MyTextAlign
	$LVcol5.Text = "SN#1"
	
	$LVcol6 = New-Object System.Windows.Forms.ColumnHeader
	$LVcol6.TextAlign = $MyTextAlign
	$LVcol6.Text = "RESULT"
	
	$LVcol7 = New-Object System.Windows.Forms.ColumnHeader
	$LVcol7.TextAlign = $MyTextAlign
	$LVcol7.Text = "START"

	$LVcol8 = New-Object System.Windows.Forms.ColumnHeader
	$LVcol8.TextAlign = $MyTextAlign
	$LVcol8.Text = "USER"
	
	$LVcol9 = New-Object System.Windows.Forms.ColumnHeader
	$LVcol9.TextAlign = $MyTextAlign
	$LVcol9.Text = "ERRORS"
	
	$LVcol10 = New-Object System.Windows.Forms.ColumnHeader
	$LVcol10.TextAlign = $MyTextAlign
	$LVcol10.Text = "FAILS"
	
	$LVcol11 = New-Object System.Windows.Forms.ColumnHeader
	$LVcol11.TextAlign = $MyTextAlign
	$LVcol11.Text = "SEQ_FILE"

	$LVcol12 = New-Object System.Windows.Forms.ColumnHeader
	$LVcol12.TextAlign = $MyTextAlign
	$LVcol12.Text = "SEQ_MD5"

	# Add the event to the ListView ColumnClick event
	$ListView.add_ColumnClick({ $listView.ListViewItemSorter = New-Object ListViewItemComparer($_.Column); UstawoKolorWierszy($ListView) })

	$ListView.Columns.AddRange([System.Windows.Forms.ColumnHeader[]](@($LVcol1, $LVcol2, $LVcol3, $LVcol4, $LVcol5, $LVcol6,  $LVcol7, $LVcol8, $LVcol9, $LVcol10, $LVcol11, $LVcol12 )))

	#write-host "Files:",($Files.gettype() | Out-String)
	#write-host "Items:",($Files | Out-String)
	
	function findMyColumn($str)
	{
		$out = ($fileContent[$nazwa] | WHERE-OBJECT { $str -in $_.keys })
		if($out)
		{
			return $out[$str]
		}
		else
		{
			return "."
		}
	}
	
	
	foreach($nazwa in ($fileContent.keys | Sort-Object ) )
	{
		#write-host "Nazwa:",($Files[$nazwa].gettype() | Out-String)
		#write-host "Items:",($Files[$nazwa] | WHERE-OBJECT { 'result' -in $_.keys } | Out-String)
				
		#write-host "Nazwa:",($Files[$nazwa].keys | Out-String)
		#wype³nanie tabeli
		$ListViewItem = New-Object System.Windows.Forms.ListViewItem([System.String[]](@($nazwa, (findMyColumn("PN\#0")), (findMyColumn("SN\#0")), (findMyColumn("PN\#1")), (findMyColumn("SN\#1")), (findMyColumn("RESULT")), (findMyColumn("START")), (findMyColumn("USER")), (findMyColumn("ERRORS")), (findMyColumn("FAILS")), (findMyColumn("SEQ_FILE")), (findMyColumn("SEQ_MD5")) )), -1) #, , , , , , , 
		#$ListViewItem.StateImageIndex = 0
		$ListView.Items.AddRange([System.Windows.Forms.ListViewItem[]](@($ListViewItem)))
		#$listView.Refresh()
	}
	
	$listView.AutoResizeColumns([System.Windows.Forms.ColumnHeaderAutoResizeStyle]::ColumnContent);
	
	UstawoKolorWierszy($ListView)
	
	$contextMenuStrip1 = New-Object System.Windows.Forms.ContextMenuStrip

	$contextMenuStrip1.Items.Add("Kopiuj ca³y wiersz").add_Click(
	{
		$item=$ListView.SelectedItems.SubItems;
		write-host ($item.Length)
		write-host ($item | Select-Object -ExpandProperty Text)
		($item | Select-Object -ExpandProperty Text) -join "`t" | Set-Clipboard
	})

	$ListView.ContextMenuStrip = $contextMenuStrip1
	
	$form.ShowDialog()
}

#https://community.spiceworks.com/topic/1982317-catch-colum-value-fromchecked-items
#$ListView.Add_MouseClick({[System.Windows.Forms.MessageBox]::Show($ListView.SelectedItems[0].Text,'Info')})
#$ListView.Add_MouseClick({Logi($ListView.SelectedItems.SubItems)})




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
$LVcol4.Text = "FP - first pass"

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


#CHECKBOX 0
$checkMe0=New-Object System.Windows.Forms.CheckBox
$checkMe0.Location=New-Object System.Drawing.Size(($Right_Row_Button+210),55)
$checkMe0.Size=New-Object System.Drawing.Size(100,30)
$checkMe0.Text="Sumuj tygodnie"
$checkMe0.TabIndex=1
$checkMe0.Checked=$false
$checkMe0.Font = $MyFont
$form.Controls.Add($checkMe0)

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
$checkMe2.Location=New-Object System.Drawing.Size(($Right_Row_Button+10),305)
$checkMe2.Size=New-Object System.Drawing.Size(150,30)
$checkMe2.Text="??"
$checkMe2.TabIndex=1
$checkMe2.Checked=$true
$checkMe2.Font = $MyFont
#$form.Controls.Add($checkMe2)

#CHECKBOX 3
$checkMe3=New-Object System.Windows.Forms.CheckBox
$checkMe3.Location=New-Object System.Drawing.Size(($Right_Row_Button+10),355)
$checkMe3.Size=New-Object System.Drawing.Size(150,30)
$checkMe3.Text="Nie dziel na daty"
$checkMe3.TabIndex=1
$checkMe3.Checked=$true
$checkMe3.Font = $MyFont
#$form.Controls.Add($checkMe3)

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

#TEXTBOX 4
$textBox4 = New-Object System.Windows.Forms.TextBox
$textBox4.Location = New-Object System.Drawing.Point(($Right_Row_Button+5),255)
$textBox4.Size = New-Object System.Drawing.Size(260,30)
$textBox4.Text=$myRegxFile
$textBox4.Font = $MyFont
#$form.Controls.Add($textBox4)


#IMAGES
[System.Windows.Forms.Application]::EnableVisualStyles();


#okno prosz¹ce o wpisanie wartoœci
function GetStringFromUser()
{
    param(
        [Parameter(Mandatory=$false, Position=0)]
        [string]$Title,
		[Parameter(Mandatory=$false, Position=1)]
		[string]$Msg,
		[Parameter(Mandatory=$false, Position=2)]
		[string]$Out
    )
	
	[void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')

	return [Microsoft.VisualBasic.Interaction]::InputBox($Msg, $Title, $Out)
}

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

#Oblicz na nowo
function ObliczPonownie
{
    param(
        [Parameter(Mandatory=$true, Position=1)]
        [hashtable]$Dict
    )
	
	$tmp = @{}
	foreach($path in $Dict.keys)
	{
		$tmp[$path] = LoadData -Dict $Dict[$path]
	}
	$tmp
}

#ponowne przeliczenie wyników z wczytanych danych
function LoadData()
{
    param(
        [Parameter(Mandatory=$true, Position=1)]
        [hashtable]$Dict
    )
	#zmienna z danymi
	$Result = @{}

	#Przetworzenie zebranych danych
	foreach($year in ($Dict.keys)) # | Sort-Object {[double]$_}))
	{	
		if($checkMe1.Checked){write-host "key{$year : ... } count value:", $Dict[$year].Length}
		
		foreach($week in ($Dict[$year].keys)) # | Sort-Object {[double]$_}))
		{
			if(-not $Dict[$year][$week]["pliki"])
			{
				continue
			}
			$fileContent = @{}
			
			$filePatternRegxKeyValue = '.*=.*'
			
			#$fileContent = $Dict[$year][$week]["pliki"].GetEnumerator() | WHERE-OBJECT { $_.Name | Select-String -Pattern $myRegxFile }
			$Dict[$year][$week]["pliki"].GetEnumerator() | WHERE-OBJECT { $_.Name | Select-String -Pattern $myRegxFile } | ForEach-Object { $fileContent.Add($_.Name, $_.Value) }

			if($checkMe1.Checked){write-host "-- N:",($Dict[$year].$week.Name | Out-String)}
			
			#PY
			$lista_pass = @()
			$lista_fail = @()
            $lista = @()

			#ile modu³ów zosta³o przetestowanych
			$lista_last_test = Zliczaj2($fileContent.GetEnumerator())
			if($checkMe1.Checked){write-host "last test:",($lista_last_test.Name | Out-String)}
			
			#Wow dzia³a
			FOREACH ($fc in $fileContent.GetEnumerator())
			{
				#write-host ($fc.Value | Out-String) 
				foreach($ff in $fc.Value)
				{
					#write-host ($ff.keys )
					if('result' -in $ff.keys)
					{
						#$List_Of_Commands.Add($Array_Object) | Out-Null
						#plik posiada log
						#write-host "+",$ff.keys, $fc.Name
						foreach($cf in $ff.keys)
						{
							#test na pass lub fail
							#write-host $cf,"=",$ff[$cf]
							
							#na pass
							if($ff[$cf] -match "pass")
							{
								#$lista_pass += $fc.Name #nazwa pliku
								$lista_pass += $fc
                                $lista += $fc
							}
							elseif($ff[$cf] -match "fail")
							{
								$lista_fail += $fc
                                $lista += $fc
							}
							else
							{
								write-host "error result",$fc.Name
							}
						}
					}
				}
			}
			
			$lista_last_pass = @($lista_last_test | WHERE-OBJECT {$lista_pass.Contains($_)} )
			if($checkMe1.Checked){write-host "Last Pass:",($lista_last_pass.Name | Out-String)}

			if($checkMe1.Checked){write-host "Pass:",($lista_pass.Name | Out-String)}
			if($checkMe1.Checked){write-host "Fail:",($lista_fail.Name | Out-String)}

			#FP
			$lista_first_pass = @($lista_pass.Name | WHERE-OBJECT { $_ -MATCH "_0.txt" })
			if($checkMe1.Checked){write-host "FP:",(($lista_pass.Name | WHERE-OBJECT { $_ -MATCH "_0.txt" }) | Out-String)}

			#FTT
			$lista_first=@(($lista.Name) | WHERE-OBJECT { $_ -MATCH "_0.txt" })
			if($checkMe1.Checked){write-host "FTT:",((($lista_pass.Name + $lista_fail.Name) | WHERE-OBJECT { $_ -MATCH "_0.txt" }) | Out-String)}
			
			$znalezione_testy = $lista_pass.Length + $lista_fail.Length
			if($checkMe1.Checked){write-host "rok:$year tydzien:$week FPY:", $lista_first_pass.Length, "/", $znalezione_testy, "PY:", $lista_pass.Length, "/", $znalezione_testy}
			
			if($Result[$year].Length -eq 0)
			{
				$Result[$year] = @{}
			}

			if(($Result[$year][$week].Length -eq 0) -and ($znalezione_testy -gt 0))
			{
				$Result[$year][$week] = @{"FPY"=$lista_first_pass.Length; "PY"=$lista_last_pass.Length ;"sum_pass"= $lista_pass.Length; "sum_test"= $znalezione_testy; "sum_moduly"= $lista_last_test.COUNT; "FTT"= $lista_first.Length; "pliki"=$fileContent}
				#write-host $Result[$year][$week]["tydzien"]
			}
			else
			{
				#write-host "Uwaga!!! Wykryto b³¹d spójnoœci testów. znalezione_testy: ",$znalezione_testy
				write-host "Uwaga!!! Wykryto nie wyœwietlane testy"
				if($checkMe1.Checked){write-host @($Dict[$year].$week | WHERE-OBJECT {-not ($lista_pass + $lista_fail).Contains($_)} )}
			}
			
		}
		if($checkMe1.Checked){write-host "koniec $year"}
		#write-host "week",$Result[$year].keys
	}
	if($checkMe1.Checked){write-host "koniec $sciezka1"}
	#write-host $Result.keys

	return ,$Result


}


#wczytanie i tworzenie danych
function GetList()
{
    param(
		[Parameter(Mandatory=$false, Position=1)]
		[string]$Sciezka1
    )
	
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
		$plik_tydzien = (Get-WeekNumber (get-date $plik.LastWriteTime -UFormat "%Y-%m-%d"))
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

        IF($checkMe0.Checked)
        {
			#wszystkie takie same rekordy sumuje razem (sumuje tygodnie razem)
		    if($Dict[$rok][$od_t].Length -eq 0)
		    {
			    $Dict[$rok][$od_t] = @($plik)
		    }
		    else
		    {
			    $Dict[$rok][$od_t] += $plik
		    }
        }
        else
        {
			#dzia³a normalnie
		    if($Dict[$rok][$plik_tydzien].Length -eq 0)
		    {
			    $Dict[$rok][$plik_tydzien] = @($plik)
		    }
		    else
		    {
			    $Dict[$rok][$plik_tydzien] += $plik
		    }
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
			if($true)
			{
			#otwarcie pliku i odczyt danych
			$fileContent = @{}
			
			#https://stackoverflow.com/questions/52709332/powershell-read-filenames-under-folder-and-read-each-file-content-to-create-menu
			#worzec wyszukiwania klucz=wartoœæ/ pomijanie lini bez takiej wartoœci np. "------"
			$filePatternRegxKeyValue = '.*=.*'
			#wype³nienie $fileContent nazwami plików jako kluczy i zawartoœci jako value
			# [Regex]::Escape - zmienia znaki ucieczki
			# ConvertFrom-StringData - zamienia na s³ownik klucz=wartoœæ ("\n\t\r \\ \..." odczytuje jako znakami ucieczki)
			#$myRegxFile, "_0.txt"
			
			#$Dict[$year].$week | ForEach-Object {$fileContent.Add($_.Name, (GET-CONTENT $_.FULLNAME -Head 10 | ForEach-Object{([Regex]::Escape($_) | Select-String -Pattern $filePatternRegxKeyValue) } | ConvertFrom-StringData))}
			
			$Dict[$year].$week | WHERE-OBJECT { $_.Name | Select-String -Pattern $myRegxFile } | ForEach-Object {$fileContent.Add($_.Name, (GET-CONTENT $_.FULLNAME -Head $ile_lini_czytac | ForEach-Object{([Regex]::Escape($_) | Select-String -Pattern $filePatternRegxKeyValue) } | ConvertFrom-StringData))}

			if($checkMe1.Checked){write-host "-- N:",($Dict[$year].$week.Name | Out-String)}
			
			#$lista_first=@(
			#write-host "--count ", $fileContent.COUNT, $fileContent.GetType() #System.Collections.Hashtable
			#write-host $fileContent.keys
			#write-host $fileContent.values  #System.Collections.Hashtable
			#write-host $fileContent[$fileContent.keys[0]].GetType() #System.Object[]
			#write-host ($fileContent[$fileContent.keys[0]].keys | Out-String) #System.Object[]
			#nazwa3_serial3_1.txt           {System.Collections.Hashtable, System.Collections.Hashtable,..
			#write-host ($fileContent[$fileContent.keys[0]][0] | Out-String) #.GetType() #System.Object[]
			#write-host $fileContent[$fileContent.keys[0]][0].GetType() #System.Object[]
			#write-host $fileContent[$fileContent.keys[0]][0][0].GetType() #System.Collections.Hashtable
			#write-host ($fileContent[$fileContent.keys[0]][0][0].keys | Out-String) #System.Collections.Hashtable
			
			
			#$tmp = @($fileContent.GetEnumerator() | WHERE-OBJECT { $_.Value } | WHERE-OBJECT{ 'result' -in $_.keys})
			#write-host ($tmp | Out-String)
			
			#PY
			$lista_pass = @()
			$lista_fail = @()
            $lista = @()

			#ile modu³ów zosta³o przetestowanych
			$lista_last_test = Zliczaj2($fileContent.GetEnumerator())
			if($checkMe1.Checked){write-host "last test:",($lista_last_test.Name | Out-String)}
			
			#Wow dzia³a
			FOREACH ($fc in $fileContent.GetEnumerator())
			{
				#write-host ($fc.Value | Out-String) 
				foreach($ff in $fc.Value)
				{
					#write-host ($ff.keys )
					if('result' -in $ff.keys)
					{
						#$List_Of_Commands.Add($Array_Object) | Out-Null
						#plik posiada log
						#write-host "+",$ff.keys, $fc.Name
						foreach($cf in $ff.keys)
						{
							#test na pass lub fail
							#write-host $cf,"=",$ff[$cf]
							
							#na pass
							if($ff[$cf] -match "pass")
							{
								#$lista_pass += $fc.Name #nazwa pliku
								$lista_pass += $fc
                                $lista += $fc
							}
							elseif($ff[$cf] -match "fail")
							{
								$lista_fail += $fc
                                $lista += $fc
							}
							else
							{
								write-host "error result",$fc.Name
							}
						}
					}
				}
			}
			
			$lista_last_pass = @($lista_last_test | WHERE-OBJECT {$lista_pass.Contains($_)} )
			if($checkMe1.Checked){write-host "Last Pass:",($lista_last_pass.Name | Out-String)}

			if($checkMe1.Checked){write-host "Pass:",($lista_pass.Name | Out-String)}
			if($checkMe1.Checked){write-host "Fail:",($lista_fail.Name | Out-String)}

			#FP
			$lista_first_pass = @($lista_pass.Name | WHERE-OBJECT { $_ -MATCH "_0.txt" })
			if($checkMe1.Checked){write-host "FP:",(($lista_pass.Name | WHERE-OBJECT { $_ -MATCH "_0.txt" }) | Out-String)}

			#FTT
			$lista_first=@(($lista.Name) | WHERE-OBJECT { $_ -MATCH "_0.txt" })
			if($checkMe1.Checked){write-host "FTT:",((($lista_pass.Name + $lista_fail.Name) | WHERE-OBJECT { $_ -MATCH "_0.txt" }) | Out-String)}

			
			
			
			
			
			}
			else
			{
			
			
			#$Dict[$year].$week | ForEach-Object {$fileContent.Add($_.Name, (GET-CONTENT $_.FULLNAME -Head 10 | ConvertFrom-StringData))}
			
			#write-host "key{$year : {$week : ... }} count value:", $Dict[$year].$week.Length
			#write-host $Dict[$year].$week

			#poprzednia wersja sprawdzania testu na pass, PY
			#$lista=@($Dict[$year].$week | WHERE-OBJECT { $A=GET-CONTENT $_.FULLNAME; $A -MATCH "FAILS=0" }  | WHERE-OBJECT { $A=GET-CONTENT $_.FULLNAME; $A -MATCH "ERRORS=0" } )
			#write-host "rok:$year tydzien:$week PY:", $lista.Length, "/", $Dict[$year].$week.Length

			#FP
			$lista_first_pass=@($Dict[$year].$week | WHERE-OBJECT { $A=GET-CONTENT $_.FULLNAME -Head $ile_lini_czytac; $A -MATCH "RESULT=PASS"} | WHERE-OBJECT { $_.Name -MATCH "_0.txt" } )

			#FTT
			$lista_first=@($Dict[$year].$week | WHERE-OBJECT { $A=GET-CONTENT $_.FULLNAME -Head $ile_lini_czytac; $A -MATCH "RESULT="} | WHERE-OBJECT { $_.Name -MATCH "_0.txt" } )
			
			#PY
			#sprawdzenie czy dany log zawiera ci¹g znaków w pierwszych 10 liniach, jesli tak to test zaliczany jako pass
			#dodaæ sprawdzanie czy plik zawieraj¹ obie linie !
			$lista_pass=@($Dict[$year].$week | WHERE-OBJECT { $A=GET-CONTENT $_.FULLNAME -Head $ile_lini_czytac; $A -MATCH "RESULT=PASS" } )
			$lista_fail=@($Dict[$year].$week | WHERE-OBJECT { $A=GET-CONTENT $_.FULLNAME -Head $ile_lini_czytac; $A -MATCH "RESULT=FAIL" } )
			
			#ile modu³ów zosta³o przetestowanych
			$lista_last_test=(Zliczaj2 $Dict[$year][$week])
			#write-host $lista_last_test
			$lista_last_pass=@($lista_last_test | WHERE-OBJECT { $A=GET-CONTENT $_.FULLNAME -Head $ile_lini_czytac; $A -MATCH "RESULT=PASS" } )
			#write-host "lista_last_pass",$lista_last_pass.Length
			}
			
			
			$znalezione_testy = $lista_pass.Length + $lista_fail.Length
			if($checkMe1.Checked){write-host "rok:$year tydzien:$week FPY:", $lista_first_pass.Length, "/", $znalezione_testy, "PY:", $lista_pass.Length, "/", $znalezione_testy}

			
			#wykrycie b³êdu w obliczeniach
			#if(($lista_pass.Length + $lista_fail.Length) -ne $Dict[$year].$week.Length)
			#{
			#	write-host "Uwaga!!! Wykryto plik ktory nie jest logiem z testow! tydzien:$week PASS + FAIL =", $lista_pass.Length, "+", $lista_fail.Length,"=",$Dict[$year].$week.Length
			#	write-host ""
			#}
			
			if($Result[$year].Length -eq 0)
			{
				$Result[$year] = @{}
			}

			if(($Result[$year][$week].Length -eq 0) -and ($znalezione_testy -gt 0))
			{
				$Result[$year][$week] = @{"FPY"=$lista_first_pass.Length; "PY"=$lista_last_pass.Length ;"sum_pass"= $lista_pass.Length; "sum_test"= $znalezione_testy; "sum_moduly"= $lista_last_test.COUNT; "FTT"= $lista_first.Length; "pliki"=$fileContent}
				#write-host $Result[$year][$week]["tydzien"]
			}
			else
			{
				#write-host "Uwaga!!! Wykryto b³¹d spójnoœci testów. znalezione_testy: ",$znalezione_testy
				write-host "Uwaga!!! Wykryto nie wyœwietlane testy"
				if($checkMe1.Checked){write-host @($Dict[$year].$week | WHERE-OBJECT {-not ($lista_pass + $lista_fail).Contains($_)} )}
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
        [Parameter(Mandatory=$false, Position=0)]
        [string]$Path,
		[Parameter(Mandatory=$false, Position=1)]
		[string]$String
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


    if ($Path -and (Test-Path $Path))
    {
        $json = Get-Content $Path -Raw
    }
    else
    {
        $json = '{}'
    }
	
	if($String)
	{
		$json = $String
	}
    else
    {
        $json = '{}'
    }
	
    $hashtable = Get-Value -value (ConvertFrom-Json $json)

    return $hashtable
}

#https://gist.github.com/marcgeld/bfacfd8d70b34fdf1db0022508b02aca
#https://powershell.org/forums/topic/compressing-data-to-gzip/
function Get-CompressedByteArray {

	[CmdletBinding()]
    Param (
	[Parameter(Mandatory,ValueFromPipeline,ValueFromPipelineByPropertyName)]
        [byte[]] $byteArray = $(Throw("-byteArray is required"))
    )
	Process {
        Write-Verbose "Get-CompressedByteArray"
       	[System.IO.MemoryStream] $output = New-Object System.IO.MemoryStream
        $gzipStream = New-Object System.IO.Compression.GzipStream $output, ([IO.Compression.CompressionMode]::Compress)
      	$gzipStream.Write( $byteArray, 0, $byteArray.Length )
        $gzipStream.Close()
        $output.Close()
        $tmp = $output.ToArray()
        Write-Output $tmp
    }
}

function Get-DecompressedByteArray {

	[CmdletBinding()]
    Param (
		[Parameter(Mandatory,ValueFromPipeline,ValueFromPipelineByPropertyName)]
        [byte[]] $byteArray = $(Throw("-byteArray is required"))
    )
	Process {
	    Write-Verbose "Get-DecompressedByteArray"
        $input = New-Object System.IO.MemoryStream( , $byteArray )
	    $output = New-Object System.IO.MemoryStream
        $gzipStream = New-Object System.IO.Compression.GzipStream $input, ([IO.Compression.CompressionMode]::Decompress)
	    $gzipStream.CopyTo( $output )
        $gzipStream.Close()
		$input.Close()
		[byte[]] $byteOutArray = $output.ToArray()
        Write-Output $byteOutArray
    }
}

#https://stackoverflow.com/questions/50654683/powershell-decompress-byte-array-takes-lot-of-time
#https://www.codeproject.com/Questions/424881/encode-and-decode-error-invalid-character-in-a-bas
#https://gallery.technet.microsoft.com/scriptcenter/GUI-popup-FileSaveDialog-813a4966
function dojson()
{
    param(
        [Parameter(Mandatory=$true, Position=1)]
        [string]$Format
    )
	
    $openDiag=New-Object System.Windows.Forms.savefiledialog
	$openDiag.initialDirectory = [System.IO.Directory]::GetCurrentDirectory()
	$openDiag.filter = "Log Files|*.$Format|All Files|*.*" 
    #otwieranie ostatnio wybrany folder
    $result=$openDiag.ShowDialog()
    if($result -eq "OK")
    {
	    #[System.Windows.Forms.MessageBox]::Show("Ustawiono: "+$openDiag.filename,'plik')
		#$Wynik | Select-Object -Property * | ConvertTo-JSON -Depth 4 | Set-Content -Path $openDiag.filename
		
		if($Format -eq "json")
		{
			#dane surowe
			$Wynik | ConvertTo-JSON -Depth 6 | Set-Content -Path $openDiag.filename
		}
		elseif($Format -eq "gz")
		{
			#dane skompresowane
			
			#tworzymy enkoder
			[System.Text.Encoding] $enc = [System.Text.Encoding]::UTF8
			#konwersja danych na json i ³¹cznie w jeden string
			$text = -join ($Wynik | ConvertTo-JSON -Depth 6)
			
			[byte[]] $encText = $enc.GetBytes( $text )
			
			$compressedByteArray = Get-CompressedByteArray -byteArray $encText
			#Write-Host "Encoded: " ( $enc.GetString( $compressedByteArray ) | Out-String )
			
			[Io.File]::WriteAllBytes($openDiag.filename, $compressedByteArray )
			#$Wynik | ForEach-OBJECT{ [pscustomobject]$_} | Export-CSV -Path "dump.csv"
		}
		elseif($Format -eq "base64")
		{
			#BASE64
			
			#tworzymy enkoder
			[System.Text.Encoding] $enc = [System.Text.Encoding]::UTF8
			#konwersja danych na json i ³¹cznie w jeden string
			$text = -join( $Wynik | ConvertTo-JSON -Depth 6 )
			#konwersja z string na bytes
			[byte[]] $toEncode2Bytes = $enc.GetBytes($text);
			#konwersja na base64
			$sReturnValue = [System.Convert]::ToBase64String($toEncode2Bytes);
			#konwersja na bytes i zapis do pliku
			[IO.File]::WriteAllBytes($openDiag.filename, $enc.GetBytes($sReturnValue))
		}
		else
		{
			write-host "Nie zanany format:",$format
		}
    }
}

#https://community.idera.com/database-tools/powershell/powertips/b/tips/posts/using-open-file-dialogs
function zjson()
{
    param(
        [Parameter(Mandatory=$true, Position=1)]
        [string]$Format
    )
	
    $openDiag=New-Object System.Windows.Forms.OpenFileDialog
	$openDiag.Multiselect = $false
	$openDiag.initialDirectory = [System.IO.Directory]::GetCurrentDirectory()
	$openDiag.filter = "Log Files|*.$Format|All Files|*.*"
    #otwieranie ostatnio wybrany folder
    $result=$openDiag.ShowDialog()
    if($result -eq "OK")
    {
	    #[System.Windows.Forms.MessageBox]::Show("Ustawiono: "+$openDiag.filename,'plik')
		#$script:Wynik = (Get-Content -Raw -Path $openDiag.filename | ConvertFrom-Json)
		#write-host ($Wynik | ConvertTo-JSON -Depth 4)
		
		if($Format -eq "json")
		{
			#dane surowe
			#$script:Wynik = Get-FromJson -Path $openDiag.filename #nie bêdzie dzia³a³o bo usun¹³em -Path
			$script:Wynik = Get-FromJson -String (Get-Content -Raw -Path $openDiag.filename)
		}
		elseif($Format -eq "gz")
		{
			#dane skompresowane
			
			#tworzymy enkoder
			[System.Text.Encoding] $enc = [System.Text.Encoding]::UTF8
			
			$inBytes = [System.IO.File]::ReadAllBytes($openDiag.filename)

			$bytes = Get-DecompressedByteArray -byteArray $inBytes
			
			$dataString = $enc.GetString( $bytes )		
			
			$script:Wynik = Get-FromJson -String $dataString
			#Write-Host "Decoded: " ( $script:Wynik | Out-String )
		}
		elseif($Format -eq "base64")
		{
			#BASE64

			#tworzymy enkoder
			[System.Text.Encoding] $enc = [System.Text.Encoding]::UTF8
			#odczyt danych i zmiana z bytes na string
			$eValue = $enc.GetString( [System.IO.File]::ReadAllBytes($openDiag.filename) )
			#konwersja z base64
			[byte[]] $encodedDataBytes = [System.Convert]::FromBase64String($eValue);
			#zmiana z bytes na string
			$sReturnValue = [System.Text.Encoding]::UTF8.GetString($encodedDataBytes);
			#odczytanie w³aœciwych informacji
			$script:Wynik = Get-FromJson -String $sReturnValue
		}
		else
		{
			write-host "Nie zanany format:",$format
		}

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
	$listView.Items.Clear()
	
	foreach($path in (Get-ChildItem $sciezka))
	{
		write-host $path,$path.LastWriteTime
		if((Get-Item $path.FULLNAME) -is [System.IO.DirectoryInfo])
		{
			#wype³nanie tabeli aktualnym statusem pracy
			$ListViewItem = New-Object System.Windows.Forms.ListViewItem([System.String[]](@($path.Name, "...", "...", "...", "...", "...", "...", "...", "...")), -1)
			#$ListViewItem.StateImageIndex = 0
			$ListView.Items.AddRange([System.Windows.Forms.ListViewItem[]](@($ListViewItem)))
			$ListView.Refresh()
			$form.Refresh()
			
			#w³aœciwe generowanie wyników
			if($checkMe1.Checked){write-host $path.FULLNAME}
			$Wynik[$path.Name] = GetList -Sciezka $path.FULLNAME
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
	New-ItemProperty -Path $regPath -Name $regMyRegxFile -Value $myRegxFile -Force | Out-Null
}

function Odswiez()
{
	zapis_konfiguracji
	
	#odczytanie zmiennych w oknach
	$testRok=$textBox1.Text
	$od_t=$textBox2.Text
	$do_t=$textBox3.Text

	#odœwierzenie listy
	$listView.Items.Clear()
	
	$label6.Text="£¹cznie folderów: " + ($Wynik.keys).COUNT
	$label6.Refresh()

	foreach($modul in ($Wynik.keys | Sort-Object ) )
	{
		if($checkMe1.Checked){write-host "modul", $modul}
		
		foreach($year in ($Wynik[$modul].keys | Sort-Object {[int]$_}))
		{
			if($checkMe1.Checked){write-host "key{$modul : {$year : ... }} count value:", $Wynik[$modul][$year].Length}
			
			foreach($week in ($Wynik[$modul][$year].keys | Sort-Object {[int]$_}))
			{
				if($checkMe1.Checked){write-host "key{$modul : {$year : {$week : ... }}} count value:", $Wynik[$modul][$year].$week.Length, $Wynik.$modul.$year.$week["FPY"],$Wynik.$modul.$year.$week["PY"],$Wynik.$modul.$year.$week["sum"]}
				
				#if($checkMe1.Checked){write-host ($Wynik.$modul.$year.$week["pliki"] | ConvertTo-JSON -Depth 2)}
				
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

				#wype³nanie tabeli
				$ListViewItem = New-Object System.Windows.Forms.ListViewItem([System.String[]](@($modul, $week, $year, $Wynik.$modul.$year.$week["FPY"], $Wynik.$modul.$year.$week["FTT"], $Wynik.$modul.$year.$week["PY"], $Wynik.$modul.$year.$week["sum_moduly"], $Wynik.$modul.$year.$week["sum_pass"], $Wynik.$modul.$year.$week["sum_test"])), -1)
				#$ListViewItem.StateImageIndex = 0
				$ListView.Items.AddRange([System.Windows.Forms.ListViewItem[]](@($ListViewItem)))
				#$listView.Refresh()
			}
		}
		
	}
	
	UstawoKolorWierszy($ListView)
	
	#$label6.Text="Wyników: " + ($ListView.Items).COUNT
	#$label6.Refresh()

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

		$label1.Text=$sciezka
		$label1.Refresh()
    }
}

$label1.Text=$sciezka
$label1.Refresh()


#pokazanie okna
$form.ShowDialog()
