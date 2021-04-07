Imports Microsoft.Office.Interop

Public Class Printanje1

    Private Sub BackgroundWorker1_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs)

    End Sub

    Private Sub Printanje1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim wd As Word.Application
        Dim wdDoc As Word.Document
        wd = New Word.Application
        wd.Visible = True
        wdDoc = wd.Documents.Add("G:\PROJEKTI 2021\Svjedodžbe Osnovne škole (Visual Basic)\Svjedodžbe Osnovne škole\templates\svjedodzba_zavrsetak_razreda.dotx") 'Putanja...'

        With wdDoc
            'Podaci o školi'
            .FormFields("skola").Result = skola.Text
            .FormFields("opcina").Result = opcina.Text

            'Matične knjige'
            .FormFields("maticna_knjiga").Result = maticna_knjiga.Text
            .FormFields("maticni_broj").Result = maticni_broj.Text

            'Završeni razred'
            If zavrsio.Text = "sedmi (VII.)" Then
                .FormFields("razred").Result = "sedmi"
                .FormFields("razred_broj").Result = "VII."
                .FormFields("razred1").Result = "sedmi"
                .FormFields("razred_broj1").Result = "VII."
                .FormFields("razred2").Result = "sedmi"
                .FormFields("razred_broj2").Result = "VII."
            ElseIf zavrsio.Text = "osmi (VIII.)" Then
                .FormFields("razred").Result = "osmi"
                .FormFields("razred_broj").Result = "VIII."
                .FormFields("razred1").Result = "osmi"
                .FormFields("razred_broj1").Result = "VIII."
                .FormFields("razred2").Result = "osmi"
                .FormFields("razred_broj2").Result = "VIII."
            End If

            'SPOL'
            If spol.Text = "M" Then
                .FormFields("spolŽ").Result = "---"
                .FormFields("spol1").Result = "in"
                .FormFields("spol2").Result = "o"
                .FormFields("spol3").Result = "ao"

                .FormFields("musko1").Result = "--"
                .FormFields("musko2").Result = "--"
                .FormFields("musko3").Result = "--"
            ElseIf spol.Text = "Ž" Then
                .FormFields("spolM").Result = "---"
                .FormFields("spol1").Result = "ka"
                .FormFields("spol2").Result = "la"
                .FormFields("spol3").Result = "la"
            End If

            'Opći podaci'
            .FormFields("imeiprezime").Result = ime.Text & " " & prezime.Text

            .FormFields("sin_kci").Result = sin_kci.Text

            .FormFields("datum_rodjenja").Result = datum_rodjenja.Text

            .FormFields("mjesto_rodjenja").Result = mjesto_rodjenja.Text

            .FormFields("opcina_rodjenja").Result = opcina_rodjenja.Text

            .FormFields("drzava_rodjenja").Result = drzava_rodjenja.Text

            .FormFields("drzavljanin").Result = drzavljanin.Text

            .FormFields("narodnost").Result = narodnost.Text

            .FormFields("godina1").Result = godina1.Text
            .FormFields("godina2").Result = godina2.Text

            .FormFields("koji_put").Result = koji_put.Text

            'OCJENE'
            'Jezik'
            If jezik1_ocjena.Text = "5" Then
                .FormFields("jezik1").Result = jezik1.Text
                .FormFields("ocjena1").Result = "odličan"
                .FormFields("ocjena_broj1").Result = "5"
            ElseIf jezik1_ocjena.Text = "4" Then
                .FormFields("jezik1").Result = jezik1.Text
                .FormFields("ocjena1").Result = "vrlo dobar"
                .FormFields("ocjena_broj1").Result = "4"
            ElseIf jezik1_ocjena.Text = "3" Then
                .FormFields("jezik1").Result = jezik1.Text
                .FormFields("ocjena1").Result = "dobar"
                .FormFields("ocjena_broj1").Result = "3"
            ElseIf jezik1_ocjena.Text = "2" Then
                .FormFields("jezik1").Result = jezik1.Text
                .FormFields("ocjena1").Result = "dovoljan"
                .FormFields("ocjena_broj1").Result = "2"
            Else
                .FormFields("jezik1").Result = "----------"
                .FormFields("ocjena1").Result = "----------"
                .FormFields("ocjena_broj1").Result = "-"
            End If

            'Likovna kultura'
            If likovna_kultura_ocjena.Text = "5" Then
                .FormFields("ocjena2").Result = "odličan"
                .FormFields("ocjena_broj2").Result = "5"
            ElseIf likovna_kultura_ocjena.Text = "4" Then
                .FormFields("ocjena2").Result = "vrlo dobar"
                .FormFields("ocjena_broj2").Result = "4"
            ElseIf likovna_kultura_ocjena.Text = "3" Then
                .FormFields("ocjena2").Result = "dobar"
                .FormFields("ocjena_broj2").Result = "3"
            ElseIf likovna_kultura_ocjena.Text = "2" Then
                .FormFields("ocjena2").Result = "dovoljan"
                .FormFields("ocjena_broj2").Result = "2"
            Else
                .FormFields("ocjena2").Result = "----------"
                .FormFields("ocjena_broj2").Result = "-"
            End If

            'Glazbena kultura'
            If glazbena_kultura_ocjena.Text = "5" Then
                .FormFields("ocjena3").Result = "odličan"
                .FormFields("ocjena_broj3").Result = "5"
            ElseIf glazbena_kultura_ocjena.Text = "4" Then
                .FormFields("ocjena3").Result = "vrlo dobar"
                .FormFields("ocjena_broj3").Result = "4"
            ElseIf glazbena_kultura_ocjena.Text = "3" Then
                .FormFields("ocjena3").Result = "dobar"
                .FormFields("ocjena_broj3").Result = "3"
            ElseIf glazbena_kultura_ocjena.Text = "2" Then
                .FormFields("ocjena3").Result = "dovoljan"
                .FormFields("ocjena_broj3").Result = "2"
            Else
                .FormFields("ocjena3").Result = "----------"
                .FormFields("ocjena_broj3").Result = "-"
            End If

            'Strani jezik1'
            If strani_jezik1_ocjena.Text = "5" Then
                .FormFields("strani_jezik1").Result = strani_jezik1.Text
                .FormFields("ocjena4").Result = "odličan"
                .FormFields("ocjena_broj4").Result = "5"
            ElseIf strani_jezik1_ocjena.Text = "4" Then
                .FormFields("strani_jezik1").Result = strani_jezik1.Text
                .FormFields("ocjena4").Result = "vrlo dobar"
                .FormFields("ocjena_broj4").Result = "4"
            ElseIf strani_jezik1_ocjena.Text = "3" Then
                .FormFields("strani_jezik1").Result = strani_jezik1.Text
                .FormFields("ocjena4").Result = "dobar"
                .FormFields("ocjena_broj4").Result = "3"
            ElseIf strani_jezik1_ocjena.Text = "2" Then
                .FormFields("strani_jezik1").Result = strani_jezik1.Text
                .FormFields("ocjena4").Result = "dovoljan"
                .FormFields("ocjena_broj4").Result = "2"
            Else
                .FormFields("strani_jezik1").Result = "----------"
                .FormFields("ocjena4").Result = "----------"
                .FormFields("ocjena_broj4").Result = "-"
            End If

            'Strani jezik2'
            If strani_jezik2_ocjena.Text = "5" Then
                .FormFields("strani_jezik2").Result = strani_jezik2.Text
                .FormFields("ocjena5").Result = "odličan"
                .FormFields("ocjena_broj5").Result = "5"
            ElseIf strani_jezik2_ocjena.Text = "4" Then
                .FormFields("strani_jezik2").Result = strani_jezik2.Text
                .FormFields("ocjena5").Result = "vrlo dobar"
                .FormFields("ocjena_broj5").Result = "4"
            ElseIf strani_jezik2_ocjena.Text = "3" Then
                .FormFields("strani_jezik2").Result = strani_jezik2.Text
                .FormFields("ocjena5").Result = "dobar"
                .FormFields("ocjena_broj5").Result = "3"
            ElseIf strani_jezik2_ocjena.Text = "2" Then
                .FormFields("strani_jezik2").Result = strani_jezik2.Text
                .FormFields("ocjena5").Result = "dovoljan"
                .FormFields("ocjena_broj5").Result = "2"
            Else
                .FormFields("strani_jezik2").Result = "----------"
                .FormFields("ocjena5").Result = "----------"
                .FormFields("ocjena_broj5").Result = "-"
            End If

            'Matematika'
            If matematika_ocjena.Text = "5" Then
                .FormFields("ocjena6").Result = "odličan"
                .FormFields("ocjena_broj6").Result = "5"
            ElseIf matematika_ocjena.Text = "4" Then
                .FormFields("ocjena6").Result = "vrlo dobar"
                .FormFields("ocjena_broj6").Result = "4"
            ElseIf matematika_ocjena.Text = "3" Then
                .FormFields("ocjena6").Result = "dobar"
                .FormFields("ocjena_broj6").Result = "3"
            ElseIf matematika_ocjena.Text = "2" Then
                .FormFields("ocjena6").Result = "dovoljan"
                .FormFields("ocjena_broj6").Result = "2"
            Else
                .FormFields("ocjena6").Result = "----------"
                .FormFields("ocjena_broj6").Result = "-"
            End If

            'Priroda'
            If priroda_ocjena.Text = "5" Then
                .FormFields("ocjena7").Result = "odličan"
                .FormFields("ocjena_broj7").Result = "5"
            ElseIf priroda_ocjena.Text = "4" Then
                .FormFields("ocjena7").Result = "vrlo dobar"
                .FormFields("ocjena_broj7").Result = "4"
            ElseIf priroda_ocjena.Text = "3" Then
                .FormFields("ocjena7").Result = "dobar"
                .FormFields("ocjena_broj7").Result = "3"
            ElseIf priroda_ocjena.Text = "2" Then
                .FormFields("ocjena7").Result = "dovoljan"
                .FormFields("ocjena_broj7").Result = "2"
            Else
                .FormFields("ocjena7").Result = "----------"
                .FormFields("ocjena_broj7").Result = "-"
            End If

            'Biologija'
            If biologija_ocjena.Text = "5" Then
                .FormFields("ocjena8").Result = "odličan"
                .FormFields("ocjena_broj8").Result = "5"
            ElseIf biologija_ocjena.Text = "4" Then
                .FormFields("ocjena8").Result = "vrlo dobar"
                .FormFields("ocjena_broj8").Result = "4"
            ElseIf biologija_ocjena.Text = "3" Then
                .FormFields("ocjena8").Result = "dobar"
                .FormFields("ocjena_broj8").Result = "3"
            ElseIf biologija_ocjena.Text = "2" Then
                .FormFields("ocjena8").Result = "dovoljan"
                .FormFields("ocjena_broj8").Result = "2"
            Else
                .FormFields("ocjena8").Result = "----------"
                .FormFields("ocjena_broj8").Result = "-"
            End If

            'Kemija'
            If kemija_ocjena.Text = "5" Then
                .FormFields("ocjena9").Result = "odličan"
                .FormFields("ocjena_broj9").Result = "5"
            ElseIf kemija_ocjena.Text = "4" Then
                .FormFields("ocjena9").Result = "vrlo dobar"
                .FormFields("ocjena_broj9").Result = "4"
            ElseIf kemija_ocjena.Text = "3" Then
                .FormFields("ocjena9").Result = "dobar"
                .FormFields("ocjena_broj9").Result = "3"
            ElseIf kemija_ocjena.Text = "2" Then
                .FormFields("ocjena9").Result = "dovoljan"
                .FormFields("ocjena_broj9").Result = "2"
            Else
                .FormFields("ocjena9").Result = "----------"
                .FormFields("ocjena_broj9").Result = "-"
            End If

            'Fizika'
            If fizika_ocjena.Text = "5" Then
                .FormFields("ocjena10").Result = "odličan"
                .FormFields("ocjena_broj10").Result = "5"
            ElseIf fizika_ocjena.Text = "4" Then
                .FormFields("ocjena10").Result = "vrlo dobar"
                .FormFields("ocjena_broj10").Result = "4"
            ElseIf fizika_ocjena.Text = "3" Then
                .FormFields("ocjena10").Result = "dobar"
                .FormFields("ocjena_broj10").Result = "3"
            ElseIf fizika_ocjena.Text = "2" Then
                .FormFields("ocjena10").Result = "dovoljan"
                .FormFields("ocjena_broj10").Result = "2"
            Else
                .FormFields("ocjena10").Result = "----------"
                .FormFields("ocjena_broj10").Result = "-"
            End If

            'Povijest'
            If povijest_ocjena.Text = "5" Then
                .FormFields("ocjena11").Result = "odličan"
                .FormFields("ocjena_broj11").Result = "5"
            ElseIf povijest_ocjena.Text = "4" Then
                .FormFields("ocjena11").Result = "vrlo dobar"
                .FormFields("ocjena_broj11").Result = "4"
            ElseIf povijest_ocjena.Text = "3" Then
                .FormFields("ocjena11").Result = "dobar"
                .FormFields("ocjena_broj11").Result = "3"
            ElseIf povijest_ocjena.Text = "2" Then
                .FormFields("ocjena11").Result = "dovoljan"
                .FormFields("ocjena_broj11").Result = "2"
            Else
                .FormFields("ocjena11").Result = "----------"
                .FormFields("ocjena_broj11").Result = "-"
            End If

            'Geografija'
            If geografija_ocjena.Text = "5" Then
                .FormFields("ocjena12").Result = "odličan"
                .FormFields("ocjena_broj12").Result = "5"
            ElseIf geografija_ocjena.Text = "4" Then
                .FormFields("ocjena12").Result = "vrlo dobar"
                .FormFields("ocjena_broj12").Result = "4"
            ElseIf geografija_ocjena.Text = "3" Then
                .FormFields("ocjena12").Result = "dobar"
                .FormFields("ocjena_broj12").Result = "3"
            ElseIf geografija_ocjena.Text = "2" Then
                .FormFields("ocjena12").Result = "dovoljan"
                .FormFields("ocjena_broj12").Result = "2"
            Else
                .FormFields("ocjena12").Result = "----------"
                .FormFields("ocjena_broj12").Result = "-"
            End If

            'Tehnička kultura'
            If tehnicka_kultura_ocjena.Text = "5" Then
                .FormFields("ocjena13").Result = "odličan"
                .FormFields("ocjena_broj13").Result = "5"
            ElseIf tehnicka_kultura_ocjena.Text = "4" Then
                .FormFields("ocjena13").Result = "vrlo dobar"
                .FormFields("ocjena_broj13").Result = "4"
            ElseIf tehnicka_kultura_ocjena.Text = "3" Then
                .FormFields("ocjena13").Result = "dobar"
                .FormFields("ocjena_broj13").Result = "3"
            ElseIf tehnicka_kultura_ocjena.Text = "2" Then
                .FormFields("ocjena13").Result = "dovoljan"
                .FormFields("ocjena_broj13").Result = "2"
            Else
                .FormFields("ocjena13").Result = "----------"
                .FormFields("ocjena_broj13").Result = "-"
            End If

            'Tjelesna i zdrvastvena kultura'
            If tjelesni_ocjena.Text = "5" Then
                .FormFields("ocjena14").Result = "odličan"
                .FormFields("ocjena_broj14").Result = "5"
            ElseIf tjelesni_ocjena.Text = "4" Then
                .FormFields("ocjena14").Result = "vrlo dobar"
                .FormFields("ocjena_broj14").Result = "4"
            ElseIf tjelesni_ocjena.Text = "3" Then
                .FormFields("ocjena14").Result = "dobar"
                .FormFields("ocjena_broj14").Result = "3"
            ElseIf tjelesni_ocjena.Text = "2" Then
                .FormFields("ocjena14").Result = "dovoljan"
                .FormFields("ocjena_broj14").Result = "2"
            Else
                .FormFields("ocjena14").Result = "----------"
                .FormFields("ocjena_broj14").Result = "-"
            End If

            'Ostali predmet 1'
            If ostali_predmeti1_ocjena.Text = "5" Then
                .FormFields("ostali_predmet1").Result = ostali_predmet1.Text
                .FormFields("ocjena15").Result = "odličan"
                .FormFields("ocjena_broj15").Result = "5"
            ElseIf ostali_predmeti1_ocjena.Text = "4" Then
                .FormFields("ostali_predmet1").Result = ostali_predmet1.Text
                .FormFields("ocjena15").Result = "vrlo dobar"
                .FormFields("ocjena_broj15").Result = "4"
            ElseIf ostali_predmeti1_ocjena.Text = "3" Then
                .FormFields("ostali_predmet1").Result = ostali_predmet1.Text
                .FormFields("ocjena15").Result = "dobar"
                .FormFields("ocjena_broj15").Result = "3"
            ElseIf ostali_predmeti1_ocjena.Text = "2" Then
                .FormFields("ostali_predmet1").Result = ostali_predmet1.Text
                .FormFields("ocjena15").Result = "dovoljan"
                .FormFields("ocjena_broj15").Result = "2"
            Else
                .FormFields("ostali_predmet1").Result = "----------"
                .FormFields("ocjena15").Result = "----------"
                .FormFields("ocjena_broj15").Result = "-"
            End If

            'Ostali predmet 2'
            If ostali_predmeti2_ocjena.Text = "5" Then
                .FormFields("ostali_predmet2").Result = ostali_predmet2.Text
                .FormFields("ocjena16").Result = "odličan"
                .FormFields("ocjena_broj16").Result = "5"
            ElseIf ostali_predmeti2_ocjena.Text = "4" Then
                .FormFields("ostali_predmet2").Result = ostali_predmet2.Text
                .FormFields("ocjena16").Result = "vrlo dobar"
                .FormFields("ocjena_broj16").Result = "4"
            ElseIf ostali_predmeti2_ocjena.Text = "3" Then
                .FormFields("ostali_predmet2").Result = ostali_predmet2.Text
                .FormFields("ocjena16").Result = "dobar"
                .FormFields("ocjena_broj16").Result = "3"
            ElseIf ostali_predmeti2_ocjena.Text = "2" Then
                .FormFields("ostali_predmet2").Result = ostali_predmet2.Text
                .FormFields("ocjena16").Result = "dovoljan"
                .FormFields("ocjena_broj16").Result = "2"
            Else
                .FormFields("ostali_predmet2").Result = "----------"
                .FormFields("ocjena16").Result = "----------"
                .FormFields("ocjena_broj16").Result = "-"
            End If

            'Ostali predmet 3'
            If ostali_predmeti3_ocjena.Text = "5" Then
                .FormFields("ostali_predmet3").Result = ostali_predmet3.Text
                .FormFields("ocjena17").Result = "odličan"
                .FormFields("ocjena_broj17").Result = "5"
            ElseIf ostali_predmeti3_ocjena.Text = "4" Then
                .FormFields("ostali_predmet3").Result = ostali_predmet3.Text
                .FormFields("ocjena17").Result = "vrlo dobar"
                .FormFields("ocjena_broj17").Result = "4"
            ElseIf ostali_predmeti3_ocjena.Text = "3" Then
                .FormFields("ostali_predmet3").Result = ostali_predmet3.Text
                .FormFields("ocjena17").Result = "dobar"
                .FormFields("ocjena_broj17").Result = "3"
            ElseIf ostali_predmeti3_ocjena.Text = "2" Then
                .FormFields("ostali_predmet3").Result = ostali_predmet3.Text
                .FormFields("ocjena17").Result = "dovoljan"
                .FormFields("ocjena_broj17").Result = "2"
            Else
                .FormFields("ostali_predmet3").Result = "----------"
                .FormFields("ocjena17").Result = "----------"
                .FormFields("ocjena_broj17").Result = "-"
            End If

            'Ostali predmet 4'
            If ostali_predmeti4_ocjena.Text = "5" Then
                .FormFields("ostali_predmet4").Result = ostali_predmet4.Text
                .FormFields("ocjena18").Result = "odličan"
                .FormFields("ocjena_broj18").Result = "5"
            ElseIf ostali_predmeti4_ocjena.Text = "4" Then
                .FormFields("ostali_predmet4").Result = ostali_predmet4.Text
                .FormFields("ocjena18").Result = "vrlo dobar"
                .FormFields("ocjena_broj18").Result = "4"
            ElseIf ostali_predmeti4_ocjena.Text = "3" Then
                .FormFields("ostali_predmet4").Result = ostali_predmet4.Text
                .FormFields("ocjena18").Result = "dobar"
                .FormFields("ocjena_broj18").Result = "3"
            ElseIf ostali_predmeti4_ocjena.Text = "2" Then
                .FormFields("ostali_predmet4").Result = ostali_predmet4.Text
                .FormFields("ocjena18").Result = "dovoljan"
                .FormFields("ocjena_broj18").Result = "2"
            Else
                .FormFields("ostali_predmet4").Result = "----------"
                .FormFields("ocjena18").Result = "----------"
                .FormFields("ocjena_broj18").Result = "-"
            End If

            'Izborni predmeti 1'
            If izborni_predmeti1_ocjena.Text = "5" Then
                .FormFields("izborni_predmeti1").Result = izborni_predmeti1.Text
                .FormFields("ocjena19").Result = "odličan"
                .FormFields("ocjena_broj19").Result = "5"
            ElseIf izborni_predmeti1_ocjena.Text = "4" Then
                .FormFields("izborni_predmeti1").Result = izborni_predmeti1.Text
                .FormFields("ocjena19").Result = "vrlo dobar"
                .FormFields("ocjena_broj19").Result = "4"
            ElseIf izborni_predmeti1_ocjena.Text = "3" Then
                .FormFields("izborni_predmeti1").Result = izborni_predmeti1.Text
                .FormFields("ocjena19").Result = "dobar"
                .FormFields("ocjena_broj19").Result = "3"
            ElseIf izborni_predmeti1_ocjena.Text = "2" Then
                .FormFields("izborni_predmeti1").Result = izborni_predmeti1.Text
                .FormFields("ocjena19").Result = "dovoljan"
                .FormFields("ocjena_broj19").Result = "2"
            Else
                .FormFields("izborni_predmeti1").Result = "----------"
                .FormFields("ocjena19").Result = "----------"
                .FormFields("ocjena_broj19").Result = "-"
            End If

            'Izborni predmeti 2'
            If izborni_predmeti2_ocjena.Text = "5" Then
                .FormFields("izborni_predmeti2").Result = izborni_predmeti2.Text
                .FormFields("ocjena20").Result = "odličan"
                .FormFields("ocjena_broj20").Result = "5"
            ElseIf izborni_predmeti2_ocjena.Text = "4" Then
                .FormFields("izborni_predmeti2").Result = izborni_predmeti2.Text
                .FormFields("ocjena20").Result = "vrlo dobar"
                .FormFields("ocjena_broj20").Result = "4"
            ElseIf izborni_predmeti2_ocjena.Text = "3" Then
                .FormFields("izborni_predmeti2").Result = izborni_predmeti2.Text
                .FormFields("ocjena20").Result = "dobar"
                .FormFields("ocjena_broj20").Result = "3"
            ElseIf izborni_predmeti2_ocjena.Text = "2" Then
                .FormFields("izborni_predmeti2").Result = izborni_predmeti2.Text
                .FormFields("ocjena20").Result = "dovoljan"
                .FormFields("ocjena_broj20").Result = "2"
            Else
                .FormFields("izborni_predmeti2").Result = "----------"
                .FormFields("ocjena20").Result = "----------"
                .FormFields("ocjena_broj20").Result = "-"
            End If

            'Izborni predmeti 3'
            If izborni_predmeti3_ocjena.Text = "5" Then
                .FormFields("izborni_predmeti3").Result = izborni_predmeti3.Text
                .FormFields("ocjena21").Result = "odličan"
                .FormFields("ocjena_broj21").Result = "5"
            ElseIf izborni_predmeti3_ocjena.Text = "4" Then
                .FormFields("izborni_predmeti3").Result = izborni_predmeti3.Text
                .FormFields("ocjena21").Result = "vrlo dobar"
                .FormFields("ocjena_broj21").Result = "4"
            ElseIf izborni_predmeti3_ocjena.Text = "3" Then
                .FormFields("izborni_predmeti3").Result = izborni_predmeti3.Text
                .FormFields("ocjena21").Result = "dobar"
                .FormFields("ocjena_broj21").Result = "3"
            ElseIf izborni_predmeti3_ocjena.Text = "2" Then
                .FormFields("izborni_predmeti3").Result = izborni_predmeti3.Text
                .FormFields("ocjena21").Result = "dovoljan"
                .FormFields("ocjena_broj21").Result = "2"
            Else
                .FormFields("izborni_predmeti3").Result = "----------"
                .FormFields("ocjena21").Result = "----------"
                .FormFields("ocjena_broj21").Result = "-"
            End If

            'Ostali podaci'
            .FormFields("aktivnost1").Result = aktivnosti1.Text
            .FormFields("aktivnost2").Result = aktivnosti2.Text
            .FormFields("aktivnost3").Result = aktivnosti3.Text

            .FormFields("vladanje").Result = vladanje.Text

            .FormFields("opravdani").Result = opravdani_izostanci.Text
            .FormFields("neopravdani").Result = neopravdani_izostanci.Text

            If uspjeh.Text = "odličan (5)" Then
                .FormFields("uspjeh").Result = "odličnim"
                .FormFields("uspjeh_broj").Result = "5"
            ElseIf uspjeh.Text = "vrlo dobar (4)" Then
                .FormFields("uspjeh").Result = "vrlo dobrim"
                .FormFields("uspjeh_broj").Result = "4"
            ElseIf uspjeh.Text = "dobar (3)" Then
                .FormFields("uspjeh").Result = "dobrim"
                .FormFields("uspjeh_broj").Result = "3"
            ElseIf uspjeh.Text = "dovoljan (2)" Then
                .FormFields("uspjeh").Result = "dovoljnim"
                .FormFields("uspjeh_broj").Result = "2"
            End If

            .FormFields("ur_broj").Result = ur_broj.Text

            .FormFields("mjesto").Result = mjesto.Text

            .FormFields("datum").Result = nadnevak1.Text & "." & nadnevak2.Text & "."
            .FormFields("datum2").Result = nadnevak3.Text

            .FormFields("razrednik").Result = razrednik.Text

            .FormFields("ravnatelj").Result = ravnatelj.Text

            .FormFields("napomena").Result = napomena.Text

        End With

        ' PUTANJA (Prepravi je..)
        wdDoc.SaveAs("G:\PROJEKTI 2021\Svjedodžbe Osnovne škole (Visual Basic)\Svjedodžbe Osnovne škole\Printane_svjedodžbe\" & ime.Text & " " & prezime.Text & " - svjedodžba.DOC")

        wdDoc.PrintPreview() 'Opens print Preview Window
        'wdDoc.PrintOut()' 'printing
        'wdDoc.SaveAs("C:\Users\Severin\Desktop\Svjedodžbe Osnovne škole\Svjedodžbe Osnovne škole\Printane_svjedodžbe" & ime.Text & " " & prezime.Text & " - svjedodžba.DOC") 'Saves the Document
        wd.Application.Quit() 'Closing Word Application
        wd = Nothing 'Releasing References to Variable
    End Sub

    Private Sub Izlaz_Click_1(sender As Object, e As EventArgs) Handles Izlaz.Click
        Me.Close()
    End Sub
End Class

'---------------------------------------------------------------------------------------------------------------------
'                       -- SEVERIN KNEŽEVIĆ --
'                   Email: knezevicseverin@gmail.com
'---------------------------------------------------------------------------------------------------------------------