Module Module1

    Sub Main()

        Dim sulkemiskasky As Boolean = False 'Bool-muuttuja ohjelman sulkemiselle
        Dim valintaKokonaislukuna As Integer 'Käyttäjän valinta menun operaatioista

        While sulkemiskasky = False 'While-silmukka pyörii kunnes käyttäjä haluaa sulkea ohjelman
            MenuRakenne() 'Menu aliohjelman kutsu
            valintaKokonaislukuna = TiedonHankinta() 'Käyttäjän valinta hankitaan aliohjelmaa hyödyntäen

            If valintaKokonaislukuna = 5 Then '5 = Ohjelma suljetaan
                sulkemiskasky = True
            Else
                Select Case valintaKokonaislukuna 'Valitaan mikä aliohjelma suoritetaan. Valitsin SELECT CASE rakenteen IF ELSEIF ELSE rakenteen sijasta helppolukuisuuden takia
                    Case 1
                        MerkkijononPituus()
                        'Merkkijonofunktio
                    Case 2
                        Ika()
                        'Ikafunktio
                    Case 3
                        Taulukko()
                        'Taulukon kasittely
                    Case 4
                        Ohmilaskuri()
                        'Resistanssilaskuri
                End Select
            End If
            Console.ReadLine()
        End While

    End Sub

    Private Sub MenuRakenne()
        'Tulostetaan menun näkymä ensin putsaten konsolin
        Console.Clear()
        Console.WriteLine("******MENU******")
        Console.WriteLine("Valinta 1: Merkkijonon kasittely")
        Console.WriteLine("Valinta 2: Ian maaritys")
        Console.WriteLine("Valinta 3: Taulukon kasittely")
        Console.WriteLine("Valinta 4: Resistanssilaskuri")
        Console.WriteLine("Valinta 5: Lopetus")

    End Sub

    Public Function TiedonHankinta() As Integer
        'Tiedonhankinta aliohjelma
        Dim valinta As Integer 'Paikallinen muuttuja, käyttäjän valinta
        Dim kayttajanAntama As String 'Käyttäjän konsoliin syöttämä valinta

        While True
            Console.WriteLine("Syota valintasi: ")
            kayttajanAntama = Console.ReadLine()

            If Integer.TryParse(kayttajanAntama, valinta) Then 'Tarkistetaan, että käyttäjän antama arvo on kokonaisluku
                If valinta < 0 Or valinta > 5 Then 'Tarkistetaan, että käyttäjän antama arvo on 1-5
                    Console.WriteLine("Vaara arvo syotetty, syota arvo valilla 1-5") 'Virheviesti
                Else
                    Exit While
                End If
            Else
                Console.WriteLine("Et syottanyt numeroa") 'Virheviesti
            End If
        End While

        Return valinta 'Palautetaan valinta mainiin

    End Function

    Private Sub Ohmilaskuri()
        Dim ppa As Object 'Valitsin käyttäjän syötettäville arvoille datatyypin Object, jotta ohjelma ei kaatuisi väärään syötettyyn arvoon
        Dim pituus As Object
        Dim ominaisvastus As Double = 0.0175
        Dim ppaOk As Boolean = False
        Dim pituusOk As Boolean = False
        Dim pyoristettyVastus As Double

        Console.Clear()
        Console.WriteLine("Tama on johtimen resistanssilaskuri")

        Do 'Silmukan alku
            Console.WriteLine(vbNewLine & "Syota johtimen poikkipinta-ala millimetreina: ")
            ppa = Console.ReadLine() 'Poikkipinta-alan syöttö                                                                                  
            Console.WriteLine(vbNewLine & "Syota johtimen pituus metreina: ")
            pituus = Console.ReadLine() 'Johtimen pituuden syöttö

            ppaOk = IsNumeric(ppa)
            pituusOk = IsNumeric(pituus) 'Tarkistus, muuttujat=numero

            If ppaOk = True And pituusOk = True Then
                pyoristettyVastus = Math.Round((pituus / ppa) * ominaisvastus, 2, MidpointRounding.ToEven) 'p*(a/l), pyöristys
                Console.WriteLine(vbNewLine & "Johtimen resistanssi on: " & pyoristettyVastus & " ohmia" & vbNewLine & "Paina ENTER jatkaaksesi")
            Else
                Console.WriteLine(vbNewLine & "Et syottanyt numeroita") 'Virheviesti
            End If
        Loop Until (ppaOk = True And pituusOk = True) 'Silmukka päättyy, kun tiedot on annettu numeroina

    End Sub

    Private Sub MerkkijononPituus()
        'Merkkijonon pituuden selvittämis-aliohjelma
        Dim syotettyMerkkijono As String 'Käyttäjän syöttämä merkkijono
        Dim merkkijononPituus As Integer 'Merkkijonon pituus

        Console.Clear()
        Console.WriteLine("Tama ohjelma laskee merkkijonon pituuden" & vbNewLine & "Ole hyva ja syota merkkijono:")

        syotettyMerkkijono = Console.ReadLine()
        merkkijononPituus = Len(syotettyMerkkijono) 'Tallennetaan merkkijonon pituus kokonaislukumuuttujaan
        Console.WriteLine("Merkkijonon pituus on: " & merkkijononPituus & vbNewLine & "Paina ENTER jatkaaksesi")

    End Sub

    Private Sub Ika()
        'Iänlaskemis aliohjelma
        Dim nykyHetki As Date = Now() 'Nykyhetki
        Dim syntymaAika 'Käyttäjän syöttämä aika
        Dim kayttajanAika As Date 'Käytettävä aika
        Dim ika As Integer 'Ikä kokonaislukuna
        Dim syotettyOikein As Boolean 'Tarkistus, että käyttäjä syöttänyt päivämäärän oikein
        Dim poistu As Boolean = False 'While-loopin ehto

        Console.Clear()
        Console.WriteLine("Ian maarittaminen" & vbNewLine & "Syota ikasi muodossa pp.kk.vvvv")

        While poistu = False

            syntymaAika = Console.ReadLine()
            syotettyOikein = IsDate(syntymaAika) 'Tarkistetaan inputin olevan Date

            If syotettyOikein = False Then
                Console.WriteLine("Et syottanyt kelvollista paivamaaraa. Yrita uudestaan") 'Virheviesti
            Else
                kayttajanAika = Convert.ToDateTime(syntymaAika) 'Muutetaan käyttäjän antama aika Date-muuttujaksi
                ika = CInt(nykyHetki.Year - kayttajanAika.Year) 'Lasketaan vuosien erotus

                If (CInt(nykyHetki.Month) < CInt(kayttajanAika.Month)) Then 'Jos käyttäjä on syntynyt kuukautena, joka on suurempi kuin nykyhetki vähennetään iästä 1
                    ika -= 1
                    poistu = True 'Loop päättyy
                ElseIf (CInt(nykyHetki.Month) = CInt(kayttajanAika.Month)) Then
                    If (CInt(nykyHetki.Day) < CInt(kayttajanAika.Day)) Then 'Jos käyttäjä on syntynyt samana kuukautena, mutta päivä on suurempi kuin nykyhetki vähennetään iästä 1
                        ika -= 1
                        poistu = True
                    Else
                        poistu = True 'Jos päivämäärä ja kuukausi on sama, ei oteta aikaa huomioon vaan pyöristetään ikä vuosien erotukseksi
                    End If
                Else
                    poistu = True 'Jos kuukausi on nykykuukautta pienenmpi, pyöristetään ikä vuosien erotukseksi
                End If
            End If
        End While

        Console.WriteLine("Ikasi on: " & ika & " vuotta!" & vbNewLine & "Paina ENTER jatkaaksesi") 'Tulostetaan ikä

    End Sub

    Private Sub Taulukko()
        Dim numerot(9) As Integer 'Julistetaan 10 merkin tietue
        Dim randomi As New Random() 'Luodaan seedi randomille
        Dim summa As Integer 'Julistetaan kokonaislukumuuttuja

        For index As Integer = 0 To 9 'Loopataan jokaiselle tietueen muistipaikalle random integer
            Dim satunnaisluku As Integer = 0
            satunnaisluku = randomi.Next(1, 11)
            numerot(index) = satunnaisluku
        Next

        Console.Clear() 'Putsataan näyttö

        Array.Sort(numerot) 'Järjestetään tietue pienuusjärjestykseen
        Array.Reverse(numerot) 'Käännetään tietue, että suurin on eka

        Console.Write("Merkkijono: ")
        For index As Integer = 0 To 9
            summa += numerot(index) 'Lasketaan summaa
            Console.Write(numerot(index) & " ")
        Next

        Console.WriteLine(vbNewLine & "Summa: " & summa & vbNewLine & "Paina ENTER jatkaaksesi")


    End Sub

End Module

