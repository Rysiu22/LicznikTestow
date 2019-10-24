#zlicza udane testy z danego tygodznia


#dane do edycji:
#czy wyœwietlaæ dodatkowe informacje, bêdzie dzia³aæ wolniej
$debug = 0

#przciski, przesuniêcie rzêdu
$Right_Row_Button = 800
$wielkosc_czcionki_okna = 10
$rozmiar_kolumn = 115




#danych poni¿ej nie edytowaæ

#czas przeznaczony na pisanie
# 2019.08.19 - 7h
# 2019.08.20 - 4h
# 2019.08.20 - 1,5h - zmiana na lepsz¹ tabele
# 2019.10.24 - 2h

#przechowuje dane pobrane z plików
$Wynik = @{}

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
$form = New-Object System.Windows.Forms.Form
$form.Text="Testy na Pass GUI wersja. 7C"
$form.Size=New-Object System.Drawing.Size(($Right_Row_Button+160),620)
#$form.topmost = $true

#1 linia
$label1=New-Object System.Windows.Forms.label
$label1.Text="..."
$label1.AutoSize=$True
$label1.Top="15"
$label1.Left="10"
$label1.Anchor="Left,Top"
$form.Controls.Add($label1)

#2 linia
$label2=New-Object System.Windows.Forms.label
$label2.Text="Rok"
$label2.AutoSize=$True
$label2.Top="255"
$label2.Left=($Right_Row_Button+10)
$label2.Anchor="Left,Top"
$form.Controls.Add($label2)

#4 linia
$label2=New-Object System.Windows.Forms.label
$label2.Text="Od tygodnia"
$label2.AutoSize=$True
$label2.Top="305"
$label2.Left=$Right_Row_Button
$label2.Anchor="Left,Top"
$form.Controls.Add($label2)

#5 linia
$label2=New-Object System.Windows.Forms.label
$label2.Text="Do tygodnia"
$label2.AutoSize=$True
$label2.Top="355"
$label2.Left=$Right_Row_Button
$label2.Anchor="Left,Top"
$form.Controls.Add($label2)

#6 linia
$label6=New-Object System.Windows.Forms.label
$label6.Text="Wyników: 0"
$label6.AutoSize=$True
$label6.Top="405"
$label6.Left=$Right_Row_Button
$label6.Anchor="Left,Top"
$form.Controls.Add($label6)

#OKNO 1
$listBox=New-Object System.Windows.Forms.Listbox
$listBox.Location = New-Object System.Drawing.Size(10,55)
$listBox.Size= New-Object System.Drawing.Size(($Right_Row_Button - 20),100)
$listbox.HorizontalScrollbar = $true;
$listBox.Font = New-Object System.Drawing.Font("Lucida Console",$wielkosc_czcionki_okna,[System.Drawing.FontStyle]::Regular)
#$form.Controls.Add($listBox)

#OKNO Z KOLUMNAMI
$listView = New-Object System.Windows.Forms.ListView
$ListView.Location = New-Object System.Drawing.Point(10, 55)
$ListView.Size = New-Object System.Drawing.Size(($Right_Row_Button - 20),500)
$ListView.View = [System.Windows.Forms.View]::Details
$ListView.FullRowSelect = $true;
$ListView.Font = New-Object System.Drawing.Font("Lucida Console",$wielkosc_czcionki_okna,[System.Drawing.FontStyle]::Regular)
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
$LVcol2.Width = $rozmiar_kolumn
$LVcol3 = New-Object System.Windows.Forms.ColumnHeader
$LVcol3.TextAlign = $MyTextAlign
$LVcol3.Text = "Rok"
$LVcol4 = New-Object System.Windows.Forms.ColumnHeader
$LVcol4.TextAlign = $MyTextAlign
$LVcol4.Text = "FPY - first pass yield"
$LVcol5 = New-Object System.Windows.Forms.ColumnHeader
$LVcol5.TextAlign = $MyTextAlign
$LVcol5.Text = "PY - pass yield"
$LVcol6 = New-Object System.Windows.Forms.ColumnHeader
$LVcol6.TextAlign = $MyTextAlign
$LVcol6.Text = "Modu³ów Suma"
$LVcol6.Width = $rozmiar_kolumn
$LVcol7 = New-Object System.Windows.Forms.ColumnHeader
$LVcol7.TextAlign = $MyTextAlign
$LVcol7.Text = "Pass Suma"
$LVcol7.Width = $rozmiar_kolumn
$LVcol8 = New-Object System.Windows.Forms.ColumnHeader
$LVcol8.TextAlign = $MyTextAlign
$LVcol8.Text = "Testów Suma"
$LVcol8.Width = $rozmiar_kolumn

$ListView.Columns.AddRange([System.Windows.Forms.ColumnHeader[]](@($LVcol1, $LVcol2, $LVcol3, $LVcol4,$LVcol5, $LVcol6, $LVcol7, $LVcol8)))

#dzia³a dobrze
#$ListViewItem = New-Object System.Windows.Forms.ListViewItem([System.String[]](@("ISA", "52", "2019", "0","1", "6", "7", "8")), -1)
#$ListViewItem.StateImageIndex = 0
#$ListView.Items.AddRange([System.Windows.Forms.ListViewItem[]](@($ListViewItem)))	

#slabo dzia³a
#$import = @("ISA", "52", "2019", "0","1", "6", "7", "8")
#ForEach($array in $import){	
#	$item = New-Object System.Windows.Forms.ListviewItem($array)
#	$listView.Items.Add($item)}

#GENERUJ
$generate=New-Object System.Windows.Forms.Button
$generate.Location=New-Object System.Drawing.Size(($Right_Row_Button+10),55)
$generate.Size=New-Object System.Drawing.Size(100,30)
$generate.Text="Generuj"
$generate.add_click({Dzialaj})
$form.Controls.Add($generate)

#FOLDER
$locate=New-Object System.Windows.Forms.Button
$locate.Location=New-Object System.Drawing.Size(($Right_Row_Button+10),105)
$locate.Size=New-Object System.Drawing.Size(100,30)
$locate.Text="Folder"
$locate.add_click({ChangeFolder})
$form.Controls.Add($locate)

#Odœwie¿
$refresh=New-Object System.Windows.Forms.Button
$refresh.Location=New-Object System.Drawing.Size(($Right_Row_Button+10),155)
$refresh.Size=New-Object System.Drawing.Size(100,30)
$refresh.Text="Odswiez"
$refresh.add_click({Odswiez})
$form.Controls.Add($refresh)

#Zapisz
$zapisz=New-Object System.Windows.Forms.Button
$zapisz.Location=New-Object System.Drawing.Size(($Right_Row_Button+10),205)
$zapisz.Size=New-Object System.Drawing.Size(100,30)
$zapisz.Text="Zapisz"
$zapisz.add_click({Zapisz})
$zapisz.Enabled = $false;
$form.Controls.Add($zapisz)


#CHECKBOX 1
$checkMe1=New-Object System.Windows.Forms.CheckBox
$checkMe1.Location=New-Object System.Drawing.Size(($Right_Row_Button+10),15)
$checkMe1.Size=New-Object System.Drawing.Size(100,30)
$checkMe1.Text="Debug"
$checkMe1.TabIndex=1
$checkMe1.Checked=$false
$form.Controls.Add($checkMe1)

#CHECKBOX 2
$checkMe2=New-Object System.Windows.Forms.CheckBox
$checkMe2.Location=New-Object System.Drawing.Size(($Right_Row_Button+10),425)
$checkMe2.Size=New-Object System.Drawing.Size(100,30)
$checkMe2.Text="Dopisuj wyniki"
$checkMe2.TabIndex=1
$checkMe2.Checked=$false
$form.Controls.Add($checkMe2)

#TEXTBOX 1
$textBox1 = New-Object System.Windows.Forms.TextBox
$textBox1.Location = New-Object System.Drawing.Point(($Right_Row_Button+70),255)
$textBox1.Size = New-Object System.Drawing.Size(40,30)
$textBox1.Text=$testRok
$form.Controls.Add($textBox1)

#TEXTBOX 2
$textBox2 = New-Object System.Windows.Forms.TextBox
$textBox2.Location = New-Object System.Drawing.Point(($Right_Row_Button+70),305)
$textBox2.Size = New-Object System.Drawing.Size(40,30)
$textBox2.Text=$od_t
$form.Controls.Add($textBox2)

#TEXTBOX 3
$textBox3 = New-Object System.Windows.Forms.TextBox
$textBox3.Location = New-Object System.Drawing.Point(($Right_Row_Button+70),355)
$textBox3.Size = New-Object System.Drawing.Size(40,30)
$textBox3.Text=$do_t
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
				$Result[$year][$week] = @{"FPY"=$lista_first_pass.Length; "PY"=$lista_last_pass.Length ;"sum_pass"= $lista_pass.Length; "sum_test"= $znalezione_testy; "sum_moduly"= $ile_mod}
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

#akcje Generowania
function Dzialaj()
{
	write-host "Start"
	zapis_konfiguracji
	
	#czy Dopisywaæ wyniki
	IF($checkMe2.Checked -ne $true)
	{
		$global:Wynik = @{}
	}
	
	foreach($path in (Get-ChildItem $sciezka))
	{
		write-host $path
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

	#odœwierzenie listy
	$listBox.Items.Clear()
	$listView.Items.Clear()
	
	foreach($modul in ($Wynik.keys))
	{
		if($checkMe1.Checked){write-host $modul}
		
		foreach($year in ($Wynik[$modul].keys | Sort-Object {[double]$_}))
		{
			#write-host "key{$modul : {$year : ... }} count value:", $Wynik[$modul][$year].Length
			
			foreach($week in ($Wynik[$modul][$year].keys | Sort-Object {[double]$_}))
			{
				if($checkMe1.Checked){write-host "key{$modul : {$year : {$week : ... }}} count value:", $Wynik[$modul][$year].$week.Length, $Wynik.$modul.$year.$week["FPY"],$Wynik.$modul.$year.$week["PY"],$Wynik.$modul.$year.$week["sum"]} 
				
				#latwe do odczytania w kodzie
				#$listBox.Items.Add("$week/$year  $modul   FPY: "+ $Wynik.$modul.$year.$week["FPY"] + "   PY: "+ $Wynik.$modul.$year.$week["PY"] + " Suma pass: " + $Wynik.$modul.$year.$week["sum_pass"] + " Suma mod: " + $Wynik.$modul.$year.$week["sum_moduly"] + " Suma_testow: " + $Wynik.$modul.$year.$week["sum_test"] )
				
				#latwe do odczytania w oknie
				#$pad = 3
				#$week_f = ([convert]::ToInt32($week,10).ToString("#00")).PadLeft(3)
				#$listBox.Items.Add("$modul $week_f/$year FPY: "+ ($Wynik.$modul.$year.$week["FPY"].ToString("#0")).PadLeft($pad) + "   PY: "+ ($Wynik.$modul.$year.$week["PY"].ToString("#0")).PadLeft($pad) + " Suma mod: " + ($Wynik.$modul.$year.$week["sum_moduly"].ToString("#0")).PadLeft($pad) + " Suma pass: " + ($Wynik.$modul.$year.$week["sum_pass"].ToString("#0")).PadLeft($pad) + " Suma_testow: " + ($Wynik.$modul.$year.$week["sum_test"].ToString("#0")).PadLeft($pad) )

				#wype³nanie tabeli
				$ListViewItem = New-Object System.Windows.Forms.ListViewItem([System.String[]](@($modul, $week, $year, $Wynik.$modul.$year.$week["FPY"], $Wynik.$modul.$year.$week["PY"], $Wynik.$modul.$year.$week["sum_moduly"], $Wynik.$modul.$year.$week["sum_pass"], $Wynik.$modul.$year.$week["sum_test"])), -1)
				#$ListViewItem.StateImageIndex = 0
				$ListView.Items.AddRange([System.Windows.Forms.ListViewItem[]](@($ListViewItem)))	
			}
		}
		
	}
	
	$label6.Text="Wyników: " + ($ListView.Items).COUNT
	$label6.Refresh()
	
	if(($Wynik.keys).COUNT -eq 0)
	{
		UpdateList
		
	}
}


function Zapisz()
{

	[System.Windows.Forms.MessageBox]::Show("Jeszcze nie dzia³a")

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
