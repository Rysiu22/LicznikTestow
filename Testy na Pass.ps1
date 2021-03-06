#zlicza udane testy z danego tygodznia

#dane do edycji:
#czy wy�wietla� dodatkowe informacje, b�dzie dzia�a� wolniej
$debug = 0

#przciski, przesuni�cie rz�du
$Right_Row_Button = 200
$Right_Row_Label = 200
$textBoxPadingRight = 170
$Right_Row_PadingY = (50 + 25)
$wielkosc_czcionki_okna = 10
$rozmiar_kolumn = 105
$wysokosc_okna = 620
$dlugosc_okna = 1150

$plik_wzorcow = "wzorce_nazw_plikow.ini"

$wzorzec_karty = "\w+\d+-\d+_B\d+W\d+S\d+_\d+\.txt"
$wzorzec_wzmacniacza = "10P000ABT\d+_AMMAWZ\d{10}_\d+\.txt"

$ile_lini_czytac = 15


#danych poni�ej nie edytowa�

#czas przeznaczony na pisanie
# 2019.08.19 - 7h
# 2019.08.20 - 4h
# 2019.08.20 - 1,5h - zmiana na lepsz� tabele
# 2019.10.24 - 2h
# 2019.10.28 - 1,5h - dodanie sortowanie po kolumnach
# 2019.11.07 - 4h - FTT, export i import
# 2019.11.08 - 4h
# 2019.11.09 - 9,5h - wczytanie kompletnych danych z nag��ka i generowanie z nich danych, filtrowanie nazw tylko przy generowaniu
# 2019.11.11 - 9h
# 2019.12.03 - 3,5h - 19:00-22:30 dodano wczytywanie wzorc�w z osobnego pliku, poprawienie kolor�w podczas sortowania, suma tygodni tylko podczas �adowania, testy z klinanym menu
# 2020.02.16 - 3,5h
# 2020.02.21 - 4h
# 2020.09.10 - 5h update do 7K
# 2020.09.13 - 3.5h naprawiono wczytywanie wzorc�w z pliku

$title = "Testy na Pass GUI, wersja 7L"

#przechowuje dane pobrane z plik�w
$Wynik = [ordered]@{}

$regPath="HKCU:\SOFTWARE\Rysiu22\TnP7K"
$name="path"
$regYear="rok"
$regPastWeek="od_tygodnia"
$regToWeek="do_tygodnia"
$regMyRegxFile="filter_plikow"
$regMyRegxDirectory="filter_katalogu"

#folder z logami
$sciezka=[System.IO.Path]::GetDirectoryName($script:MyInvocation.MyCommand.Path) #aktualna �cie�ka
$testRok=get-date -UFormat "%Y"
$od_t="1"
$do_t="52"
$myRegxFile=".*"
$myRegxDirectory=".*"

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
	$myRegxDirectory=(Get-Item -Path $regPath).GetValue($regMyRegxDirectory)
	
}
ELSE
{
	[System.Windows.Forms.MessageBox]::Show("Pierwsze uruchomienie! Ustaw poprawnie wszystkie pola i folder. Nast�pnie wci�nij Generuj",'Info')
}

#POBIERA AKTUALN� DATE
$dzien=get-date -UFormat "%Y-%m-%d"
function Get-WeekNumber([datetime]$DateTime = (Get-Date)) 
{
	$ci = [System.Globalization.CultureInfo]::CurrentCulture
	($ci.Calendar.GetWeekOfYear($DateTime,$ci.DateTimeFormat.CalendarWeekRule,$ci.DateTimeFormat.FirstDayOfWeek)).ToString()
}

#Tworzenie okna programu
Add-Type -AssemblyName System.Windows.Forms
$form = New-Object System.Windows.Forms.Form
$form.Text=$title
$form.Size=New-Object System.Drawing.Size($dlugosc_okna, $wysokosc_okna)
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
	#'1 Od�wie�'={Odswiez}; 
	'2 �aduj'={Dzialaj}; 
	'3 Zmie� Folder'={ChangeFolder}
	'4 Przelicz ponownie'={$Wynik = ObliczPonownie $Wynik; Odswiez};
	'5 Czy�� Tabele'={$script:Wynik=@{}; Odswiez}
}

ForEach ($GroupUserKey in ($DropDownGUsers2Dict.keys | Sort-Object)) {
	#Write-Host $GroupUserKey, $DropDownGUsers2Dict[$GroupUserKey]
	$GroupValue = New-Object System.Windows.Forms.ToolStripMenuItem
	$GroupValue.Text = $GroupUserKey.Substring($GroupUserKey.IndexOf(" "))
	# name the control
	$Groupvalue.Name = $GroupUserKey
	$UserGMenu2.DropDownItems.Add($GroupValue) | Out-Null
	# use name to identify control
	$GroupValue.Add_Click( $DropDownGUsers2Dict[$GroupUserKey] )
}

$DropDownGUsers3Dict=@{
	'01 Folder W�asny' = {
		$tmp=GetStringFromUser "Info" "Podaj w�asny wzorzec" $script:myRegxDirectory; 
		if($tmp){$script:myRegxDirectory=$tmp}; 
		$label9.Text="wzorzec folderu: ",$myRegxDirectory; 
		$label9.Refresh();
		};
	'02 Foldery Wszystkie: .*' = {
		$script:myRegxDirectory=".*"; 
		$label9.Text="wzorzec folderu: ",$myRegxDirectory; 
		$label9.Refresh();
		};
	'51 Modu�y w�asny' = {
		$tmp=GetStringFromUser "Info" "Podaj w�asny wzorzec" $script:myRegxFile; 
		if($tmp){$script:myRegxFile=$tmp}; 
		$label8.Text="wzorzec pliku: ",$myRegxFile; 
		$label8.Refresh();
		$Wynik = ObliczPonownie $Wynik; Odswiez;
		};
	'52 Modu�y Wszystkie: .*' = {
		$script:myRegxFile=".*"; 
		$label8.Text="wzorzec pliku: ",$myRegxFile; 
		$label8.Refresh();
		$Wynik = ObliczPonownie $Wynik; Odswiez;
		};
}

#czyta zawarto�� pliku z wzorcami i dodaje wzorce do menu
If([System.IO.File]::Exists($plik_wzorcow))
{
	$wzorce = @(GET-CONTENT $plik_wzorcow | ForEach-Object{[Regex]::Escape($_) | Select-String -Pattern ".+=.+" } | ConvertFrom-StringData)
	#write-host ($wzorce | Out-String) -ForegroundColor yellow
	
	# nr aby  menu by�o uporz�dkowane
	$i_folder = 2
	$i_karta = 52
	ForEach($key in $wzorce.keys)
	{
		#write-host ($key | Out-String) -ForegroundColor red
		#write-host ($wzorce.$key | Out-String) -ForegroundColor red
		
		function setMyLabel8($arg)
		{
			$script:myRegxFile=$arg
			$label8.Text="wzorzec: ",$script:myRegxFile
			$label8.Refresh();
			$Wynik = ObliczPonownie $Wynik; Odswiez;
			#write-host ($arg | Out-String) -ForegroundColor yellow
		}
		function setMyLabel9($arg)
		{
			$script:myRegxDirectory=$arg;
			$label9.Text="wzorzec folderu: ",$script:myRegxDirectory;
			$label9.Refresh();
			#write-host ($arg | Out-String) -ForegroundColor yellow
		}

		if( ($key -MATCH "folder") -or ($key -MATCH "foldery") -or ($key -MATCH "katalog") -or ($key -MATCH "katalogi") )
		{
			$tmpu = ((++$i_folder).ToString('00') +"  "+($key -replace "\\ ", " "))
			#write-host ($tmpu | Out-String) -ForegroundColor yellow
			$DropDownGUsers3Dict.Add($tmpu, [scriptblock]::Create('setMyLabel9 "'+$wzorce.$key+'"') )
		}
		else
		{
			$tmpu = ((++$i_karta).ToString('00') +"  "+($key -replace "\\ ", " "))
			#write-host ($tmpu | Out-String) -ForegroundColor yellow
			$DropDownGUsers3Dict.Add($tmpu, [scriptblock]::Create('setMyLabel8 "'+$wzorce.$key+'"') )
		}
		#write-host "+",($DropDownGUsers3Dict.keys | Sort-Object | Select-Object -Last 1)
	}
	write-host "Wczytano plik z wzorcami"
}
ELSE
{
	write-host "Nie znaleziono pliku z wzorcami"
}

ForEach ($GroupUserKey in ($DropDownGUsers3Dict.keys | Sort-Object)) {
	#Write-Host $GroupUserKey, $DropDownGUsers2Dict[$GroupUserKey]
	$GroupValue = New-Object System.Windows.Forms.ToolStripMenuItem
	$GroupValue.Text = $GroupUserKey.Substring($GroupUserKey.IndexOf(" "))
	# name the control
	$Groupvalue.Name = $GroupUserKey
	$UserGMenu3.DropDownItems.Add($GroupValue) | Out-Null
	# use name to identify control
	#write-host ($GroupUserKey | Out-String) -ForegroundColor yellow
	#write-host ($DropDownGUsers3Dict.$GroupUserKey | Out-String) -ForegroundColor yellow
	$GroupValue.Add_Click( $DropDownGUsers3Dict[$GroupUserKey] )
}

$DropDownGUsers4Dict=@{
	'1 Zawsze na wieszku'={$form.topmost = -not $form.topmost};
}

ForEach ($GroupUserKey in ($DropDownGUsers4Dict.keys | Sort-Object)) {
	#Write-Host $GroupUserKey, $DropDownGUsers1Dict[$GroupUserKey]
	$GroupValue = New-Object System.Windows.Forms.ToolStripMenuItem
	$GroupValue.Text = $GroupUserKey.Substring($GroupUserKey.IndexOf(" "))
	# name the control
	$Groupvalue.Name = $GroupUserKey
	$UserGMenu4.DropDownItems.Add($GroupValue) | Out-Null
	# use name to identify control
	$GroupValue.Add_Click( $DropDownGUsers4Dict[$GroupUserKey] )
}

$DropDownGUsers5Dict=@{
	'1 Strona projektu'={Start-Process "https://github.com/Rysiu22/LicznikTestow"};
	'2 Pobierz Aktualn� wersje'={
		if( $script:MyInvocation.MyCommand.Path ){Invoke-WebRequest -Uri "https://raw.githubusercontent.com/Rysiu22/LicznikTestow/master/Testy%20na%20Pass.ps1" -OutFile $script:MyInvocation.MyCommand.Path};
		$form.Close();
		};
	
	
}

ForEach ($GroupUserKey in ($DropDownGUsers5Dict.keys | Sort-Object)) {
	#Write-Host $GroupUserKey, $DropDownGUsers1Dict[$GroupUserKey]
	$GroupValue = New-Object System.Windows.Forms.ToolStripMenuItem
	$GroupValue.Text = $GroupUserKey.Substring($GroupUserKey.IndexOf(" "))
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
$label2.Top=(35 + $Right_Row_PadingY)
$label2.Left=($form.Size.Width - $Right_Row_Label)
$label2.Anchor="Left,Top"
$label2.Font = $MyFont
$form.Controls.Add($label2)

#3 linia
$label3=New-Object System.Windows.Forms.label
$label3.Text="DATA: " + (get-date -UFormat "%Y-%m-%d")
$label3.AutoSize=$True
$label3.Top=(245 + $Right_Row_PadingY)
$label3.Left=($form.Size.Width - $Right_Row_Label)
$label3.Anchor="Left,Top"
$label3.Font = $MyFont
$form.Controls.Add($label3)

#4 linia
$label4=New-Object System.Windows.Forms.label
$label4.Text="Od tygodnia"
$label4.AutoSize=$True
$label4.Top=(85 + $Right_Row_PadingY)
$label4.Left=($form.Size.Width - $Right_Row_Label)
$label4.Anchor="Left,Top"
$label4.Font = $MyFont
$form.Controls.Add($label4)

#5 linia
$label5=New-Object System.Windows.Forms.label
$label5.Text="Do tygodnia"
$label5.AutoSize=$True
$label5.Top=(135 + $Right_Row_PadingY)
$label5.Left=($form.Size.Width - $Right_Row_Label)
$label5.Anchor="Left,Top"
$label5.Font = $MyFont
$form.Controls.Add($label5)

#6 linia
$label6=New-Object System.Windows.Forms.label
$label6.Text="��cznie folder�w: 0"
$label6.AutoSize=$True
$label6.Top=(205 + $Right_Row_PadingY)
$label6.Left=($form.Size.Width - $Right_Row_Label)
$label6.Anchor="Left,Top"
$label6.Font = $MyFont
$form.Controls.Add($label6)

#7 linia
$label7=New-Object System.Windows.Forms.label
$label7.Text="Aktualny tydzie�: " + (Get-WeekNumber)
$label7.AutoSize=$True
$label7.Top=(225 + $Right_Row_PadingY)
$label7.Left=($form.Size.Width - $Right_Row_Label)
$label7.Anchor="Left,Top"
$label7.Font = $MyFont
$form.Controls.Add($label7)

#8 linia
$label8=New-Object System.Windows.Forms.label
$label8.Text="wzorzec pliku: ",$myRegxFile
#$label8.AutoSize=$True
$label8.Size = New-Object System.Drawing.Size($dlugosc_okna, 26);
$label8.Top="55"
$label8.Left=10
$label8.Anchor="Left,Top"
$label8.Font = $MyFont
$form.Controls.Add($label8)

#9 linia
$label9=New-Object System.Windows.Forms.label
$label9.Text="wzorzec folderu: ",$myRegxDirectory
$label9.AutoSize=$True
$label9.Top=(55 + 25)
$label9.Left=10
$label9.Anchor="Left,Top"
$label9.Font = $MyFont
$form.Controls.Add($label9)

$form.add_Resize({
	$label2.Left=($form.Size.Width - $Right_Row_Label)
	$label3.Left=($form.Size.Width - $Right_Row_Label)
	$label4.Left=($form.Size.Width - $Right_Row_Label)
	$label5.Left=($form.Size.Width - $Right_Row_Label)
	$label6.Left=($form.Size.Width - $Right_Row_Label)
	$label7.Left=($form.Size.Width - $Right_Row_Label)
})

#OKNO Z KOLUMNAMI
$listView = New-Object System.Windows.Forms.ListView
$ListView.Location = New-Object System.Drawing.Point(10, (35 + $Right_Row_PadingY))
$ListView.Size = New-Object System.Drawing.Size(($form.Size.Width - $Right_Row_Label - 30),($form.Size.Height - 100 - $Right_Row_PadingY))
$ListView.View = [System.Windows.Forms.View]::Details
$ListView.FullRowSelect = $true;
$form.add_Resize({
	$ListView.Size = New-Object System.Drawing.Size(($form.Size.Width - $Right_Row_Label - 30),($form.Size.Height - 100 - $Right_Row_PadingY))
})



$contextMenuStrip1 = New-Object System.Windows.Forms.ContextMenuStrip
$contextMenuStrip1.Items.Add("Kopiuj FP, FTT").add_Click(
{
	$item=$ListView.SelectedItems.SubItems;
	$tmp=$Wynik[$item[0].Text][$item[2].Text][$item[1].Text]
	if(($tmp["FPY"]).ToString() -ne $item[3].Text -or ($tmp["FTT"]).ToString() -ne $item[4].Text)
	{
		write-host("Uwaga skopiowana warto�� mo�e by� nie poprawna")
	}
	#(($tmp["FPY"]).ToString() + "	" + ($tmp["FTT"]).ToString() | Set-Clipboard)
	($item[3].Text + "	" + $item[4].Text | Set-Clipboard)
})

$contextMenuStrip1.Items.Add("Kopiuj PY, Modu��w Suma").add_Click(
{
	$item=$ListView.SelectedItems.SubItems;
	$tmp=$Wynik[$item[0].Text][$item[2].Text][$item[1].Text]
	if(($tmp["PY"]).ToString() -ne $item[5].Text -or ($tmp["sum_moduly"]).ToString() -ne $item[6].Text)
	{
		write-host("Uwaga skopiowana warto�� mo�e by� nie poprawna")
	}
	#(($tmp["PY"]).ToString() + "	" + ($tmp["sum_moduly"]).ToString() | Set-Clipboard)
	($item[5].Text + "	" + $item[6].Text | Set-Clipboard)
})
$contextMenuStrip1.Items.Add("Pliki *").add_Click(
{
	#write-host("Zaznaczone:")
	#write-host($ListView.SelectedItems.SubItems.COUNT)
	#write-host($ListView.SelectedItems.SubItems)

	Logi($ListView.SelectedItems.SubItems)
})
$contextMenuStrip1.Items.Add("Pliki tylko ostatnie *").add_Click(
{
	#write-host("Zaznaczone:")
	#write-host($ListView.SelectedItems.SubItems.COUNT)
	#write-host($ListView.SelectedItems.SubItems)

	Logi_last($ListView.SelectedItems.SubItems)
})
$contextMenuStrip1.Items.Add("Kopiuj ca�y wiersz *").add_Click(
{
	$item=$ListView.SelectedItems.SubItems;
	$out = ""
	$ListView.Columns | ForEach-Object { $out += ($_.Text + "`t") } # Nazwy kolumn
	$out += "`r`n"
	for($i=0; $i -lt $item.Length; $i++)
	{
		if(($i % $ListView.Columns.COUNT) -eq 0 -and $i -gt 0 )
		{
			$out += "`r`n"
		}
		$out += ($item[$i].Text + "`t")
	}
	#($item | Select-Object -ExpandProperty Text) -join "`t" | Set-Clipboard
	$out | Set-Clipboard
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

$ff = @"
    using System;
    using System.IO;
	using System.Text.RegularExpressions;
    
	public class MyParse
    {
        public static DateTime ParseDateTime(string value)
        {
            //'2017-02-27c14:21.36 [3571046496,24]'
            // date 8 dig, time 6 dig, 10+2 dig (odejmuj�c '2082841200,00' uzyskamy poprawn� ilo�� sekund)

            string[] numbers = Regex.Split(value, @"\D+");

            DateTime date = DateTime.MinValue;
            if (numbers.Length >= 6)
            {
                date = new DateTime(int.Parse(numbers[0]), int.Parse(numbers[1]), int.Parse(numbers[2]), int.Parse(numbers[3]), int.Parse(numbers[4]), int.Parse(numbers[5]));
            }
            else
            {
                throw new FormatException();
            }

            return date;
        }
	}
"@

Add-Type -TypeDefinition $ff -Language CSharp


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
	$fileContent = @{}

	# $Wynik[$item[0].Text][$item[2].Text][$item[1].Text]["pliki"].GetEnumerator() | WHERE-OBJECT { $_.Name | Select-String -Pattern $myRegxFile } | ForEach-Object { $fileContent.Add($_.Name, $_.Value) }

	for($i=0; $i -lt $item.COUNT; $i+=9) #$ListView.Columns.COUNT
	{
		$Wynik[$item[$i].Text][$item[$i+2].Text][$item[$i+1].Text]["pliki"].GetEnumerator() | WHERE-OBJECT { $_.Name | Select-String -Pattern $myRegxFile } | ForEach-Object { $fileContent.Add($_.Name, $_.Value) }
	}

	if($checkMe1.Checked){write-host ($fileContent | ConvertTo-JSON -Depth 2)}

    Logi_go($fileContent)
	
}

function Logi_last($item)
{
	$fileContent = @{}
	# write-host $item.COUNT
	# $Wynik[$item[0].Text][$item[2].Text][$item[1].Text]["pliki"].GetEnumerator() | WHERE-OBJECT { $_.Name | Select-String -Pattern $myRegxFile } | ForEach-Object { $fileContent.Add($_.Name, $_.Value) }
	for($i=0; $i -lt $item.COUNT; $i+=9) #$ListView.Columns.COUNT
	{
		Zliczaj2($Wynik[$item[$i].Text][$item[$i+2].Text][$item[$i+1].Text]["pliki"].GetEnumerator()) | WHERE-OBJECT { $_.Name | Select-String -Pattern $myRegxFile } | ForEach-Object { $fileContent.Add($_.Name, $_.Value) }
	}

	if($checkMe1.Checked){write-host ($fileContent | ConvertTo-JSON -Depth 2)}
	
    Logi_go($fileContent)
	
}

function Logi_go($fileContent)
{

	if(! $item)
	{
		return
	}
	$startLoad = Get-Date

    #write-host $last
    #Write-Host $item


	#Tworzenie okna programu
	Add-Type -AssemblyName System.Windows.Forms
	$form = New-Object System.Windows.Forms.Form
	$form.Text=$item[0].Text
	$form.Size=New-Object System.Drawing.Size($dlugosc_okna, $wysokosc_okna)
	$form.StartPosition='CenterScreen'

	#OKNO Z KOLUMNAMI
	$listView = New-Object System.Windows.Forms.ListView
	$ListView.Location = New-Object System.Drawing.Point(10, 15)
	$ListView.Size = New-Object System.Drawing.Size(($form.Size.Width - 50),($form.Size.Height - 70))
	$ListView.View = [System.Windows.Forms.View]::Details
	$ListView.FullRowSelect = $true;
	$ListView.Font = $MyFont
	$form.add_Resize({
		$ListView.Size = New-Object System.Drawing.Size(($form.Size.Width - 50),($form.Size.Height - 70))
	})
	$form.Controls.Add($ListView)

	$MyTextAlign = [System.Windows.Forms.HorizontalAlignment]::Left;
	
	#"PN#0","SN#0","PN#1","SN#1","RESULT","START","STOP","USER","ERRORS","FAILS","SEQ_FILE","SEQ_MD5"

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
	$LVcol8.Text = "STOP"

	$LVcol9 = New-Object System.Windows.Forms.ColumnHeader
	$LVcol9.TextAlign = $MyTextAlign
	$LVcol9.Text = "USER"
	
	$LVcol10 = New-Object System.Windows.Forms.ColumnHeader
	$LVcol10.TextAlign = $MyTextAlign
	$LVcol10.Text = "ERRORS"
	
	$LVcol11 = New-Object System.Windows.Forms.ColumnHeader
	$LVcol11.TextAlign = $MyTextAlign
	$LVcol11.Text = "FAILS"
	
	$LVcol12 = New-Object System.Windows.Forms.ColumnHeader
	$LVcol12.TextAlign = $MyTextAlign
	$LVcol12.Text = "SEQ_FILE"

	$LVcol13 = New-Object System.Windows.Forms.ColumnHeader
	$LVcol13.TextAlign = $MyTextAlign
	$LVcol13.Text = "SEQ_MD5"

	# Add the event to the ListView ColumnClick event
	$ListView.add_ColumnClick({ $listView.ListViewItemSorter = New-Object ListViewItemComparer($_.Column); UstawoKolorWierszy($ListView) })

	$ListView.Columns.AddRange([System.Windows.Forms.ColumnHeader[]](@($LVcol1, $LVcol2, $LVcol3, $LVcol4, $LVcol5, $LVcol6,  $LVcol7, $LVcol8, $LVcol9, $LVcol10, $LVcol11, $LVcol12, $LVcol13 )))

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
			return " "
		}
	}
	
	
	foreach($nazwa in ($fileContent.keys | Sort-Object ) )
	{
		#write-host "Nazwa:",($Files[$nazwa].gettype() | Out-String)
		#write-host "Items:",($Files[$nazwa] | WHERE-OBJECT { 'result' -in $_.keys } | Out-String)
				
		#write-host "Nazwa:",($Files[$nazwa].keys | Out-String)
		#wype�nanie tabeli
		$ListViewItem = New-Object System.Windows.Forms.ListViewItem([System.String[]](@($nazwa, (findMyColumn("PN\#0")), (findMyColumn("SN\#0")), (findMyColumn("PN\#1")), (findMyColumn("SN\#1")), (findMyColumn("RESULT")), (findMyColumn("START")), (findMyColumn("STOP")), (findMyColumn("USER")), (findMyColumn("ERRORS")), (findMyColumn("FAILS")), (findMyColumn("SEQ_FILE")), (findMyColumn("SEQ_MD5")) )), -1)
		#$ListViewItem.StateImageIndex = 0
		$ListView.Items.AddRange([System.Windows.Forms.ListViewItem[]](@($ListViewItem)))
		#$listView.Refresh()
	}
	
	$listView.AutoResizeColumns([System.Windows.Forms.ColumnHeaderAutoResizeStyle]::ColumnContent);
	
	UstawoKolorWierszy($ListView)
	
	$contextMenuStrip1 = New-Object System.Windows.Forms.ContextMenuStrip

	$contextMenuStrip1.Items.Add("Kopiuj ca�y wiersz *").add_Click(
	{
		$item=$ListView.SelectedItems.SubItems;
		$out = ""
		$ListView.Columns | ForEach-Object { $out += ($_.Text + "`t") } # Nazwy kolumn
		$out += "`r`n"
		for($i=0; $i -lt $item.Length; $i++)
		{
			if(($i % $ListView.Columns.COUNT) -eq 0 -and $i -gt 0 )
			{
				$out += "`r`n"
			}
			$out += ($item[$i].Text + "`t")
		}
		#($item | Select-Object -ExpandProperty Text) -join "`t" | Set-Clipboard
		$out | Set-Clipboard
	})

	$contextMenuStrip1.Items.Add("oblicz czasy *").add_Click(
	{
		$item=$ListView.SelectedItems.SubItems;
		$out = ""
		$out_timespan = New-Object Collections.Generic.List[TimeSpan]
		#write-host("count:")
		#write-host($ListView.Columns.COUNT)
		for($i=0; $i -lt ($item.Length - 1); $i+=$ListView.Columns.COUNT)
		{
			#write-host("czasy:")
			#write-host($i)
			#write-host($item[$i+6].Text)
			#write-host($item[$i+7].Text)
			$new_timespan = ( NEW-TIMESPAN �Start ([MyParse]::ParseDateTime($item[$i+6].Text)) �End ([MyParse]::ParseDateTime($item[$i+7].Text)) )
			$out_timespan.Add($new_timespan)
			#$out += $new_timespan
			#$out += "`r`n"
		}
		#$out += "Ilo�� test�w: " + ($item.Length / $ListView.Columns.COUNT).ToString() + ". "
		#$out += "Suma czasu: $out_timespan"
		#$out_timespan | Measure TotalSecs -Average -Sum -MAx -Min|ft *
		$out += "Max: "
		$out += ($out_timespan | Measure-Object -Maximum ).Maximum
		
		$out += "Min: "
		$out += ($out_timespan | Measure-Object -Minimum ).Minimum
		
		#write-host( [TimeSpan][Int](($out_timespan | Measure-Object -Average -Property Ticks ).Average/100 ) )
		#write-host( $out_timespan | Measure TotalSecs -Average -Sum -MAx -Min )
		
		
		#$out += "Avg: " + ($out_timespan / $ListView.Columns.COUNT).toString() +"`r`n"
		#($item | Select-Object -ExpandProperty Text) -join "`t" | Set-Clipboard
		#$out | Set-Clipboard
		GetStringFromUser "Info" "Obliczono czasy" $out;
	})
	
	$contextMenuStrip1.Items.Add("kopiuj czasy *").add_Click(
	{
		$item=$ListView.SelectedItems.SubItems;
		$out = ""
		$out_timespan = New-Object Collections.Generic.List[TimeSpan]
		#write-host("count:")
		#write-host($ListView.Columns.COUNT)
		for($i=0; $i -lt ($item.Length - 1); $i+=$ListView.Columns.COUNT)
		{
			#write-host("czasy:")
			#write-host($i)
			#write-host($item[$i+6].Text)
			#write-host($item[$i+7].Text)
			$new_timespan = ( NEW-TIMESPAN �Start ([MyParse]::ParseDateTime($item[$i+6].Text)) �End ([MyParse]::ParseDateTime($item[$i+7].Text)) )
			$out_timespan.Add($new_timespan)
			$out += $new_timespan
			$out += "`r`n"
		}
		#$out += "Ilo�� test�w: " + ($item.Length / $ListView.Columns.COUNT).ToString() + ". "
		#$out += "Suma czasu: $out_timespan"
		#$out_timespan | Measure TotalSecs -Average -Sum -MAx -Min|ft *
		$out += "`r`nMax: "
		$out += ($out_timespan | Measure-Object -Maximum ).Maximum
		
		$out += "`r`nMin: "
		$out += ($out_timespan | Measure-Object -Minimum ).Minimum
		
		#write-host( [TimeSpan][Int](($out_timespan | Measure-Object -Average -Property Ticks ).Average/100 ) )
		#write-host( $out_timespan | Measure TotalSecs -Average -Sum -MAx -Min )
		
		
		#$out += "Avg: " + ($out_timespan / $ListView.Columns.COUNT).toString() +"`r`n"
		#($item | Select-Object -ExpandProperty Text) -join "`t" | Set-Clipboard
		$out | Set-Clipboard
		#GetStringFromUser "Info" "Obliczono czasy" $out;
	})


	$contextMenuStrip1.Items.Add("kopiuj czasy bez : *").add_Click(
	{
		$item=$ListView.SelectedItems.SubItems;
		$out = ""
		$out_timespan = New-Object Collections.Generic.List[TimeSpan]
		#write-host("count:")
		#write-host($ListView.Columns.COUNT)
		for($i=0; $i -lt ($item.Length - 1); $i+=$ListView.Columns.COUNT)
		{
			#write-host("czasy:")
			#write-host($i)
			#write-host($item[$i+6].Text)
			#write-host($item[$i+7].Text)
			$new_timespan = ( NEW-TIMESPAN �Start ([MyParse]::ParseDateTime($item[$i+6].Text)) �End ([MyParse]::ParseDateTime($item[$i+7].Text)) )
			$out_timespan.Add($new_timespan)
			$out += ($new_timespan -replace ':','')
			$out += "`r`n"
		}
		#$out += "Ilo�� test�w: " + ($item.Length / $ListView.Columns.COUNT).ToString() + ". "
		#$out += "Suma czasu: $out_timespan"
		#$out_timespan | Measure TotalSecs -Average -Sum -MAx -Min|ft *

		$out_tmp = "Max: `t"
		$out_tmp += ($out_timespan | Measure-Object -Maximum ).Maximum
		
		$out_tmp += "`r`nMin: `t"
		$out_tmp += ($out_timespan | Measure-Object -Minimum ).Minimum

		#$out_tmp += "`r`nSuma czasu: `t$out_timespan"
		
		$out_tmp += "`r`nMediana: `t"
		$out_tmp += ("=MEDIANA(A4:A" + (4 + $out_timespan.COUNT).ToString() + ")`r`n" )
		
		#write-host( [TimeSpan][Int](($out_timespan | Measure-Object -Average -Property Ticks ).Average/100 ) )
		#write-host( $out_timespan | Measure TotalSecs -Average -Sum -MAx -Min )
		
		
		#$out += "Avg: " + ($out_timespan / $ListView.Columns.COUNT).toString() +"`r`n"
		#($item | Select-Object -ExpandProperty Text) -join "`t" | Set-Clipboard
		($out_tmp + $out) | Set-Clipboard
		#GetStringFromUser "Info" "Obliczono czasy" $out;
	})

	$ListView.ContextMenuStrip = $contextMenuStrip1

	write-host "Koniec pliki", (NEW-TIMESPAN �Start $startLoad �End (Get-Date))

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
$LVcol2.Text = "Tydzie�"

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
$LVcol7.Text = "Modu��w Suma"
$LVcol7.Width = $rozmiar_kolumn

$LVcol8 = New-Object System.Windows.Forms.ColumnHeader
$LVcol8.TextAlign = $MyTextAlign
$LVcol8.Text = "Pass Suma"
$LVcol8.Width = 10 # $rozmiar_kolumn

$LVcol9 = New-Object System.Windows.Forms.ColumnHeader
$LVcol9.TextAlign = $MyTextAlign
$LVcol9.Text = "Test�w Suma"
$LVcol9.Width = 10 # $rozmiar_kolumn

$LVcol10 = New-Object System.Windows.Forms.ColumnHeader
$LVcol10.TextAlign = $MyTextAlign
$LVcol10.Text = "FP %"
#$LVcol10.Width = $rozmiar_kolumn

$LVcol11 = New-Object System.Windows.Forms.ColumnHeader
$LVcol11.TextAlign = $MyTextAlign
$LVcol11.Text = "PY %"
#$LVcol11.Width = $rozmiar_kolumn

$LVcol12 = New-Object System.Windows.Forms.ColumnHeader
$LVcol12.TextAlign = $MyTextAlign
$LVcol12.Text = "T/M - testy/modu�y"
#$LVcol12.Width = $rozmiar_kolumn


$ListView.Columns.AddRange([System.Windows.Forms.ColumnHeader[]](@($LVcol1, $LVcol2, $LVcol3, $LVcol4, $LVcol5, $LVcol6, $LVcol7, $LVcol8, $LVcol9, $LVcol10, $LVcol11, $LVcol12)))

#dzia�a dobrze
#$ListViewItem = New-Object System.Windows.Forms.ListViewItem([System.String[]](@("ISA", "52", "2019", "0","1", "6", "7", "8")), -1)
#$ListViewItem.StateImageIndex = 0
#$ListView.Items.AddRange([System.Windows.Forms.ListViewItem[]](@($ListViewItem)))	

#slabo dzia�a
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
$checkMe0.Location=New-Object System.Drawing.Size(($form.Size.Width - $Right_Row_Button),(285 + $Right_Row_PadingY))
$checkMe0.Size=New-Object System.Drawing.Size(100,30)
$checkMe0.Text="Sumuj tygodnie"
$checkMe0.TabIndex=1
$checkMe0.Checked=$false
$checkMe0.Font = $MyFont
$form.add_Resize({
	$checkMe0.Location = New-Object System.Drawing.Size(($form.Size.Width - $Right_Row_Button),(285 + $Right_Row_PadingY))
})
$form.Controls.Add($checkMe0)

#CHECKBOX 1
$checkMe1=New-Object System.Windows.Forms.CheckBox
$checkMe1.Location=New-Object System.Drawing.Size(($form.Size.Width - $Right_Row_Button),(325 + $Right_Row_PadingY))
$checkMe1.Size=New-Object System.Drawing.Size(100,30)
$checkMe1.Text="Debug"
$checkMe1.TabIndex=1
$checkMe1.Checked=$false
$checkMe1.Font = $MyFont
$form.add_Resize({
	$checkMe1.Location=New-Object System.Drawing.Size(($form.Size.Width - $Right_Row_Button),(325 + $Right_Row_PadingY))
})
#$form.Controls.Add($checkMe1)

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

#CHECKBOX 4
$checkMe4=New-Object System.Windows.Forms.CheckBox
$checkMe4.Location=New-Object System.Drawing.Size(($form.Size.Width - $Right_Row_Button),(325 + $Right_Row_PadingY))
$checkMe4.Size=New-Object System.Drawing.Size(100,30)
$checkMe4.Text="Post�p"
$checkMe4.TabIndex=1
$checkMe4.Checked=$true
$checkMe4.Font = $MyFont
$form.add_Resize({
	$checkMe4.Location=New-Object System.Drawing.Size(($form.Size.Width - $Right_Row_Button),(325 + $Right_Row_PadingY))
})
$form.Controls.Add($checkMe4)


#TEXTBOX 1
$textBox1 = New-Object System.Windows.Forms.TextBox
$textBox1.Location = New-Object System.Drawing.Point(($form.Size.Width - $textBoxPadingRight),(55 + $Right_Row_PadingY))
$textBox1.Size = New-Object System.Drawing.Size(40,30)
$textBox1.Text=$testRok
$textBox1.Font = $MyFont
$form.add_Resize({
	$textBox1.Location=New-Object System.Drawing.Size(($form.Size.Width - $textBoxPadingRight),(55 + $Right_Row_PadingY))
})
$form.Controls.Add($textBox1)

#TEXTBOX 2
$textBox2 = New-Object System.Windows.Forms.TextBox
$textBox2.Location = New-Object System.Drawing.Point(($form.Size.Width - $textBoxPadingRight),(105 + $Right_Row_PadingY))
$textBox2.Size = New-Object System.Drawing.Size(40,30)
$textBox2.Text=$od_t
$textBox2.Font = $MyFont
$form.add_Resize({
	$textBox2.Location=New-Object System.Drawing.Size(($form.Size.Width - $textBoxPadingRight),(105 + $Right_Row_PadingY))
})
$form.Controls.Add($textBox2)

#TEXTBOX 3
$textBox3 = New-Object System.Windows.Forms.TextBox
$textBox3.Location = New-Object System.Drawing.Point(($form.Size.Width - $textBoxPadingRight),(155 + $Right_Row_PadingY))
$textBox3.Size = New-Object System.Drawing.Size(40,30)
$textBox3.Text=$do_t
$textBox3.Font = $MyFont
$form.add_Resize({
	$textBox3.Location=New-Object System.Drawing.Size(($form.Size.Width - $textBoxPadingRight),(155 + $Right_Row_PadingY))
})
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


#okno prosz�ce o wpisanie warto�ci
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

#zwraca ilo�c modu��w /del
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


#zwraca liste plik�w z ostatnimi testami modu��w
function Zliczaj2($co)
{
	#tworzy slownik[sn_modulu] = (max_tygodniowy_numer_testu, dane_do_pliku)
	# ilo�c modu��w to: slownik.keys.COUNT
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

#ponowne przeliczenie wynik�w z wczytanych danych
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

			#ile modu��w zosta�o przetestowanych
			$lista_last_test = Zliczaj2($fileContent.GetEnumerator())
			if($checkMe1.Checked){write-host "last test:",($lista_last_test.Name | Out-String)}
			
			#Wow dzia�a
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
				#write-host "Uwaga!!! Wykryto b��d sp�jno�ci test�w. znalezione_testy: ",$znalezione_testy
				write-host "Uwaga!!! Wykryto nie wy�wietlane testy", "rok:", $year, "tydzien", $week
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
	
	#Progress parameter
	$Progress = @{
		Activity = 'Czytanie plikow:'
		CurrentOperation = "$sciezka1"
		Status = 'Wczytywanie'
		PercentComplete = 0
	}

	if($checkMe4.Checked){Write-Progress @Progress}
	$i_Progress = 0
	$i_next = 0

	#pobranie posortowanych plik�w
	$pliki = (Get-ChildItem $sciezka1 *.TXT) # | sort LastWriteTime

	#wype�nianie zmiennej danymi na temat log�w {rok:{tydzien: [dane] }
	foreach($plik in $pliki)
	{
		$i_Progress++
		
		if( $i_Progress -ge $i_next -and $checkMe4.Checked)
		{
			$i_next += [int]($pliki.Count * 0.01)
			$progress.PercentComplete = [int]($i_Progress / $pliki.Count *100)
			$progress.Status = ('Szukanie plikow: ' + $progress.PercentComplete.ToString() + '%')
			Write-Progress @Progress
		}

		$rok = (get-date $plik.LastWriteTime -UFormat %Y)
		$plik_tydzien = (Get-WeekNumber (get-date $plik.LastWriteTime -UFormat "%Y-%m-%d"))
		#if($debug){write-host $plik, (get-date $plik.LastWriteTime -UFormat "%Y.%V")}
		
		#je�li rok = 0 to pomija wszelkie restrykcjie czasowe
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
			#dzia�a normalnie
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

	#wy�wietlenie struktury rok i tydzie�
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

	$progress.Status = 'Wczytywanie'
	$progress.PercentComplete = 0
	Write-Progress @Progress
		
	#Przetworzenie zebranych danych
	foreach($year in ($Dict.keys | Sort-Object {[double]$_}))
	{	
		$i_Progress = 0
		$i_next = 0
		
		if($checkMe1.Checked){write-host "key{$year : ... } count value:", $Dict[$year].Length}
		
		foreach($week in ($Dict[$year].keys)) # | Sort-Object {[double]$_}))
		{
			$i_Progress++
			
			if( $i_Progress -ge $i_next -and $checkMe4.Checked)
			{
				$i_next += [int]($Dict[$year].Count * 0.01)
				$progress.PercentComplete = [int]($i_Progress / $Dict[$year].Count *100)
				$progress.Status = ('Rok ' + $year + '. Przetwarzanie danych: ' + $progress.PercentComplete.ToString() + '%')
				Write-Progress @Progress
			}

			if($true)
			{
			#otwarcie pliku i odczyt danych
			$fileContent = @{}
			
			#https://stackoverflow.com/questions/52709332/powershell-read-filenames-under-folder-and-read-each-file-content-to-create-menu
			#worzec wyszukiwania klucz=warto��/ pomijanie lini bez takiej warto�ci np. "------"
			$filePatternRegxKeyValue = '.*=.*'
			#wype�nienie $fileContent nazwami plik�w jako kluczy i zawarto�ci jako value
			# [Regex]::Escape - zmienia znaki ucieczki
			# ConvertFrom-StringData - zamienia na s�ownik klucz=warto�� ("\n\t\r \\ \..." odczytuje jako znakami ucieczki)
			#$myRegxFile, "_0.txt"
			
			#$Dict[$year].$week | ForEach-Object {$fileContent.Add($_.Name, (GET-CONTENT $_.FULLNAME -Head 10 | ForEach-Object{([Regex]::Escape($_) | Select-String -Pattern $filePatternRegxKeyValue) } | ConvertFrom-StringData))}
			
			#skip empty file
			$Dict[$year].$week | WHERE-OBJECT { $_.Name | Select-String -Pattern $myRegxFile } | ForEach-Object {if($checkMe4.Checked){$progress.CurrentOperation = "Odczyt: " + $_.Name; Write-Progress @Progress}; $fileContent.Add($_.Name, (GET-CONTENT $_.FULLNAME -Head $ile_lini_czytac | ? {$_.trim() -ne "" } | ForEach-Object{([Regex]::Escape($_) | Select-String -Pattern $filePatternRegxKeyValue) } | ConvertFrom-StringData))}
			#wysypuje si� na pustych plikach
			#$Dict[$year].$week | WHERE-OBJECT { $_.Name | Select-String -Pattern $myRegxFile } | ForEach-Object {$progress.CurrentOperation = "Odczyt: " + $_.Name; Write-Progress @Progress; $fileContent.Add($_.Name, (GET-CONTENT $_.FULLNAME -Head $ile_lini_czytac | ForEach-Object{([Regex]::Escape($_) | Select-String -Pattern $filePatternRegxKeyValue) } | ConvertFrom-StringData))}

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

			#ile modu��w zosta�o przetestowanych
			$lista_last_test = Zliczaj2($fileContent.GetEnumerator())
			if($checkMe1.Checked){write-host "last test:",($lista_last_test.Name | Out-String)}
			
			#Wow dzia�a
			FOREACH ($fc in $fileContent.GetEnumerator())
			{
				#write-host ($fc.Value | Out-String)
				if($checkMe4.Checked)
				{
					$progress.CurrentOperation = "Przetwarzanie: " + $fc.Name
					Write-Progress @Progress
				}

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
			#write-host ($lista_last_test | ConvertTo-JSON -Depth 4 | Out-String)
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
			#sprawdzenie czy dany log zawiera ci�g znak�w w pierwszych 10 liniach, jesli tak to test zaliczany jako pass
			#doda� sprawdzanie czy plik zawieraj� obie linie !
			$lista_pass=@($Dict[$year].$week | WHERE-OBJECT { $A=GET-CONTENT $_.FULLNAME -Head $ile_lini_czytac; $A -MATCH "RESULT=PASS" } )
			$lista_fail=@($Dict[$year].$week | WHERE-OBJECT { $A=GET-CONTENT $_.FULLNAME -Head $ile_lini_czytac; $A -MATCH "RESULT=FAIL" } )
			
			#ile modu��w zosta�o przetestowanych
			$lista_last_test=(Zliczaj2 $Dict[$year][$week])
			#write-host $lista_last_test
			$lista_last_pass=@($lista_last_test | WHERE-OBJECT { $A=GET-CONTENT $_.FULLNAME -Head $ile_lini_czytac; $A -MATCH "RESULT=PASS" } )
			#write-host "lista_last_pass",$lista_last_pass.Length
			}
			
			
			$znalezione_testy = $lista_pass.Length + $lista_fail.Length
			if($checkMe1.Checked){write-host "rok:$year tydzien:$week FPY:", $lista_first_pass.Length, "/", $znalezione_testy, "PY:", $lista_pass.Length, "/", $znalezione_testy}

			
			#wykrycie b��du w obliczeniach
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
				#write-host "Uwaga!!! Wykryto b��d sp�jno�ci test�w. znalezione_testy: ",$znalezione_testy
				write-host "Uwaga!!! Wykryto nie wy�wietlane testy", "rok:", $year, "tydzien", $week
				if($checkMe1.Checked){write-host @($Dict[$year].$week | WHERE-OBJECT {-not ($lista_pass + $lista_fail).Contains($_)} )}
			}
			
		}
		if($checkMe1.Checked){write-host "koniec $year"}
		#write-host "week",$Result[$year].keys
	}
	if($checkMe1.Checked){write-host "koniec $sciezka1"}
	#write-host $Result.keys
	
	Write-Progress -Completed close

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
			#konwersja danych na json i ��cznie w jeden string
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
			#konwersja danych na json i ��cznie w jeden string
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
			#$script:Wynik = Get-FromJson -Path $openDiag.filename #nie b�dzie dzia�a�o bo usun��em -Path
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
			#odczytanie w�a�ciwych informacji
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
	$startLoad = Get-Date
	zapis_konfiguracji
	
	#zerowanie zmiennej
	$script:Wynik = [ordered]@{}
	$listView.Items.Clear()

	$testRok=$textBox1.Text
	$od_t=$textBox2.Text
	$do_t=$textBox3.Text
	
	foreach($path in (Get-ChildItem $sciezka | WHERE-OBJECT { $_.Name | Select-String -Pattern $myRegxDirectory }))
	{
		$rok = (get-date $path.LastWriteTime -UFormat %Y)
		$plik_tydzien = (Get-WeekNumber (get-date $path.LastWriteTime -UFormat "%Y-%m-%d"))
		#if($debug){write-host $path, (get-date $path.LastWriteTime -UFormat "%Y.%V")}

		write-host $path,$path.LastWriteTime

		if((Get-Item $path.FULLNAME) -is [System.IO.DirectoryInfo])
		{
			#wype�nanie tabeli aktualnym statusem pracy
			$ListViewItem = New-Object System.Windows.Forms.ListViewItem([System.String[]](@($path.Name, "...", "...", "...", "...", "...", "...", "...", "...")), -1)
			#$ListViewItem.StateImageIndex = 0
			$ListView.Items.AddRange([System.Windows.Forms.ListViewItem[]](@($ListViewItem)))
			$ListView.Refresh()
			$form.Refresh()
			
			#w�a�ciwe generowanie wynik�w
			if($checkMe1.Checked){write-host $path.FULLNAME}
			$Wynik[$path.Name] = GetList -Sciezka $path.FULLNAME
		}
	}
	
	Odswiez
	write-host "Koniec", (NEW-TIMESPAN �Start $startLoad �End (Get-Date))
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
	New-ItemProperty -Path $regPath -Name $regMyRegxDirectory -Value $myRegxDirectory -Force | Out-Null
}

function Odswiez()
{
	zapis_konfiguracji
	
	#odczytanie zmiennych w oknach
	$testRok=$textBox1.Text
	$od_t=$textBox2.Text
	$do_t=$textBox3.Text

	#od�wierzenie listy
	$listView.Items.Clear()
	
	$label6.Text="��cznie folder�w: " + ($Wynik.keys).COUNT
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
				
				#je�li rok = 0 to pomija wszelkie restrykcjie czasowe
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

				#wype�nanie tabeli
				$fpy = "--"
				if( ($Wynik.$modul.$year.$week["FPY"] -ne 0) -and ($Wynik.$modul.$year.$week["FTT"] -gt 9) ) # -and (($Wynik.$modul.$year.$week["FTT"]*2) -gt $Wynik.$modul.$year.$week["sum_moduly"]))
				{
					$fpy = [math]::Round(($Wynik.$modul.$year.$week["FPY"] * 100 / $Wynik.$modul.$year.$week["FTT"]))
				}
				$py = "--"
				if( ($Wynik.$modul.$year.$week["PY"] -ne 0) -and ($Wynik.$modul.$year.$week["sum_moduly"] -gt 9) )
				{
					$py = [math]::Round(($Wynik.$modul.$year.$week["PY"] * 100 / $Wynik.$modul.$year.$week["sum_moduly"]))
				}
				$e = [math]::Round(($Wynik.$modul.$year.$week["sum_test"] / $Wynik.$modul.$year.$week["sum_moduly"]), 2)
				$ListViewItem = New-Object System.Windows.Forms.ListViewItem([System.String[]](@($modul, $week, $year, $Wynik.$modul.$year.$week["FPY"], $Wynik.$modul.$year.$week["FTT"], $Wynik.$modul.$year.$week["PY"], $Wynik.$modul.$year.$week["sum_moduly"], $Wynik.$modul.$year.$week["sum_pass"], $Wynik.$modul.$year.$week["sum_test"], $fpy, $py, $e)), -1)
				#$ListViewItem.StateImageIndex = 0
				$ListView.Items.AddRange([System.Windows.Forms.ListViewItem[]](@($ListViewItem)))
				#$listView.Refresh()
			}
		}
		
	}
	
	UstawoKolorWierszy($ListView)
	
	#$label6.Text="Wynik�w: " + ($ListView.Items).COUNT
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
