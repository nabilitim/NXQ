
Imports MySql.Data.MySqlClient
Imports System.Math
Imports System.Globalization
Public Class Form1


    ' Id connexion pour l'ancienne base.
    'Dim SQLcon As MySqlConnection = New MySqlConnection("SERVER=srv-bdd1;Port=3306;DATABASE=test;UID=user;PASSWORD=Cpw5pU2Q5esbAFDb")

    ' Id connexion pour la nouvelle base.
    Dim SQLcon As MySqlConnection = New MySqlConnection("SERVER=maintenance;Port=3306;DATABASE=suivife;UID=suivife_RW;PASSWORD=cavBWuSnUG33jNcv")


    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        If Button1.Text = "Connection" Then

            Try
                If SQLcon.State = ConnectionState.Closed Then
                    SQLcon.Open()
                    Button1.Text = "Se déconnecter de la Base de Donnée"
                    Button1.BackColor = Color.Green
                Else
                    SQLcon.Close()
                    Button1.Text = "Connection"
                    Button1.BackColor = Color.Yellow
                End If
            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try
        ElseIf Button1.Text = "Se déconnecter de la Base de Donnée" Then
            Try
                If SQLcon.State = ConnectionState.Open Then
                    SQLcon.Close()
                    Button1.Text = "Connection"
                    Button1.BackColor = Color.Yellow
                Else
                    SQLcon.Open()
                    Button1.Text = "Se déconnecter de la Base de Donnée"
                    Button1.BackColor = Color.Green
                End If
            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try
        End If

    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        TextBox1.Text = TextBox11.Text
        TextBox11.Select()
        CultureInfo.CurrentCulture = New CultureInfo("en-GB")
        If TextBox11.Text = "" Then
            MessageBox.Show("Veuillez entrée une valeur!")

        ElseIf TextBox1.Text < 7.0 Then
            Button5.BackColor = Color.Red
            TextBox13.Text = TextBox13.Text + " " + "Point N°1 hors spec = " & TextBox1.Text

        ElseIf TextBox1.Text > 9.0 Then
            Button5.BackColor = Color.Red
            TextBox13.Text = TextBox13.Text + " " + "Point N°1 hors spec = " & TextBox1.Text

        Else
            Button5.BackColor = Color.Green
        End If
        TextBox11.Text = ""

    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        TextBox2.Text = TextBox11.Text
        CultureInfo.CurrentCulture = New CultureInfo("en-GB")
        TextBox11.Select()
        If TextBox11.Text = "" Then
            MessageBox.Show("Veuillez entrée une valeur!")
        ElseIf TextBox2.Text < 7.0 Then
            Button4.BackColor = Color.Red
            TextBox13.Text = TextBox13.Text + " " + "Point N°2 hors spec = " & TextBox2.Text

        ElseIf TextBox2.Text > 9.0 Then
            Button4.BackColor = Color.Red
            TextBox13.Text = TextBox13.Text + " " + "Point N°2 hors spec = " & TextBox2.Text

        Else
            Button4.BackColor = Color.Green
        End If

        TextBox11.Text = ""
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        TextBox3.Text = TextBox11.Text
        CultureInfo.CurrentCulture = New CultureInfo("en-GB")
        TextBox11.Select()
        If TextBox11.Text = "" Then
            MessageBox.Show("Veuillez entrée une valeur!")
        ElseIf TextBox3.Text < 7.0 Then
            Button6.BackColor = Color.Red
            TextBox13.Text = TextBox13.Text + " " + "Point N°3 hors spec = " & TextBox3.Text

        ElseIf TextBox3.Text > 9.0 Then
            Button6.BackColor = Color.Red
            TextBox13.Text = TextBox13.Text + " " + "Point N°3 hors spec = " & TextBox3.Text
        Else
            Button6.BackColor = Color.Green
        End If

        TextBox11.Text = ""
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        TextBox4.Text = TextBox11.Text
        CultureInfo.CurrentCulture = New CultureInfo("en-GB")
        TextBox11.Select()
        If TextBox11.Text = "" Then
            MessageBox.Show("Veuillez entrée une valeur!")
        ElseIf TextBox4.Text < 7.0 Then
            Button7.BackColor = Color.Red
            TextBox13.Text = TextBox13.Text + " " + "Point N°4 hors spec = " & TextBox4.Text

        ElseIf TextBox4.Text > 9.0 Then
            Button7.BackColor = Color.Red
            TextBox13.Text = TextBox13.Text + " " + "Point N°4 hors spec = " & TextBox4.Text
        Else
            Button7.BackColor = Color.Green
        End If

        TextBox11.Text = ""
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        TextBox5.Text = TextBox11.Text
        CultureInfo.CurrentCulture = New CultureInfo("en-GB")
        TextBox11.Select()
        If TextBox11.Text = "" Then
            MessageBox.Show("Veuillez entrée une valeur!")
        ElseIf TextBox5.Text < 7.0 Then
            Button8.BackColor = Color.Red
            TextBox13.Text = TextBox13.Text + " " + "Point N°5 hors spec = " & TextBox5.Text
        ElseIf TextBox5.Text > 9.0 Then
            Button8.BackColor = Color.Red
            TextBox13.Text = TextBox13.Text + " " + "Point N°5 hors spec = " & TextBox5.Text
        Else
            Button8.BackColor = Color.Green
        End If

        TextBox11.Text = ""
    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        TextBox6.Text = TextBox11.Text
        CultureInfo.CurrentCulture = New CultureInfo("en-GB")
        TextBox11.Select()
        If TextBox11.Text = "" Then
            MessageBox.Show("Veuillez entrée une valeur!")
        ElseIf TextBox6.Text < 7.0 Then
            Button9.BackColor = Color.Red
            TextBox13.Text = TextBox13.Text + " " + "Point N°6 hors spec = " & TextBox6.Text
        ElseIf TextBox6.Text > 9.0 Then
            Button9.BackColor = Color.Red
            TextBox13.Text = TextBox13.Text + " " + "Point N°6 hors spec = " & TextBox6.Text
        Else
            Button9.BackColor = Color.Green
        End If

        TextBox11.Text = ""
    End Sub

    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
        TextBox7.Text = TextBox11.Text
        CultureInfo.CurrentCulture = New CultureInfo("en-GB")
        TextBox11.Select()
        If TextBox11.Text = "" Then
            MessageBox.Show("Veuillez entrée une valeur!")
        ElseIf TextBox7.Text < 7.0 Then
            Button10.BackColor = Color.Red
            TextBox13.Text = TextBox13.Text + " " + "Point N°7 hors spec = " & TextBox7.Text
        ElseIf TextBox7.Text > 9.0 Then
            Button10.BackColor = Color.Red
            TextBox13.Text = TextBox13.Text + " " + "Point N°7 hors spec = " & TextBox7.Text
        Else
            Button10.BackColor = Color.Green
        End If

        TextBox11.Text = ""
    End Sub

    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
        TextBox8.Text = TextBox11.Text
        CultureInfo.CurrentCulture = New CultureInfo("en-GB")
        TextBox11.Select()
        If TextBox11.Text = "" Then
            MessageBox.Show("Veuillez entrée une valeur!")
        ElseIf TextBox8.Text < 7.0 Then
            Button11.BackColor = Color.Red
            TextBox13.Text = TextBox13.Text + " " + "Point N°8 hors spec = " & TextBox8.Text
        ElseIf TextBox8.Text > 9.0 Then
            Button11.BackColor = Color.Red
            TextBox13.Text = TextBox13.Text + " " + "Point N°8 hors spec = " & TextBox8.Text
        Else
            Button11.BackColor = Color.Green
        End If

        TextBox11.Text = ""
    End Sub

    Private Sub Button12_Click(sender As Object, e As EventArgs) Handles Button12.Click
        TextBox9.Text = TextBox11.Text
        CultureInfo.CurrentCulture = New CultureInfo("en-GB")
        TextBox11.Select()
        If TextBox11.Text = "" Then
            MessageBox.Show("Veuillez entrée une valeur!")
        ElseIf TextBox9.Text < 7.0 Then
            Button12.BackColor = Color.Red
            TextBox13.Text = TextBox13.Text + " " + "Point N°9 hors spec = " & TextBox9.Text
        ElseIf TextBox9.Text > 9.0 Then
            Button12.BackColor = Color.Red
            TextBox13.Text = TextBox13.Text + " " + "Point N°9 hors spec = " & TextBox9.Text
        Else
            Button12.BackColor = Color.Green
        End If

        TextBox11.Text = ""
    End Sub
    Private Sub Label11_Click(sender As Object, e As EventArgs) Handles Label11.Click
        TextBox10.Text = DateTimePicker1.Value
    End Sub

    Private Sub Label2_Click(sender As Object, e As EventArgs) Handles Label2.Click
        TextBox1.Text = " "

    End Sub

    Private Sub Label3_Click(sender As Object, e As EventArgs) Handles Label3.Click
        TextBox2.Text = " "
    End Sub

    Private Sub Label4_Click(sender As Object, e As EventArgs) Handles Label4.Click
        TextBox3.Text = " "
    End Sub

    Private Sub Label5_Click(sender As Object, e As EventArgs) Handles Label5.Click
        TextBox4.Text = " "
    End Sub

    Private Sub Label6_Click(sender As Object, e As EventArgs) Handles Label6.Click
        TextBox5.Text = " "
    End Sub

    Private Sub Label7_Click(sender As Object, e As EventArgs) Handles Label7.Click
        TextBox6.Text = " "
    End Sub

    Private Sub Label8_Click(sender As Object, e As EventArgs) Handles Label8.Click
        TextBox7.Text = " "
    End Sub

    Private Sub Label9_Click(sender As Object, e As EventArgs) Handles Label9.Click
        TextBox8.Text = " "
    End Sub

    Private Sub Label10_Click(sender As Object, e As EventArgs) Handles Label10.Click
        TextBox9.Text = " "
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If SQLcon.State = ConnectionState.Open Then

            If TextBox10.Text.Trim() = "" Then
                MessageBox.Show("Entrez une Date faire un click sur Date")
            ElseIf TextBox1.Text.Trim() = "" Then
                MessageBox.Show("Entrez LE POINT N°1")
            ElseIf TextBox2.Text.Trim() = "" Then
                MessageBox.Show("Entrez LE POINT N°2")
            ElseIf TextBox3.Text.Trim() = "" Then
                MessageBox.Show("Entrez LE POINT N°3")
            ElseIf TextBox4.Text.Trim() = "" Then
                MessageBox.Show("Entrez LE POINT N°4")
            ElseIf TextBox5.Text.Trim() = "" Then
                MessageBox.Show("Entrez LE POINT N°5")
            ElseIf TextBox6.Text.Trim() = "" Then
                MessageBox.Show("Entrez LE POINT N°6")
            ElseIf TextBox7.Text.Trim() = "" Then
                MessageBox.Show("Entrez LE POINT N°7")
            ElseIf TextBox8.Text.Trim() = "" Then
                MessageBox.Show("Entrez LE POINT N°8")
            ElseIf TextBox9.Text.Trim() = "" Then
                MessageBox.Show("Entrez LE POINT N°9")
            ElseIf ComboBox1.Text.Trim() = "" Then
                MessageBox.Show("Entrez un Visa")
            ElseIf TextBox13.Text.Trim() = "" Then
                MessageBox.Show("Entrez un Commentaire")


            Else

                Dim cmd As New MySqlCommand("INSERT INTO `nxq_pm_lampe` (Date, point_n1, point_n2, point_n3, point_n4, point_n5, point_n6, point_n7, point_n8, point_n9,average,unif, visa, commentaire) VALUES(@Date,@point_n1,@point_n2,@point_n3,@point_n4,@point_n5,@point_n6,@point_n7,@point_n8,@point_n9,@average,@unif,@visa,@commentaire)", SQLcon)
                cmd.Parameters.AddWithValue("@Date", DateTime.Parse(TextBox10.Text).ToString("yyyy-MM-dd HH:mm:ss"))
                cmd.Parameters.AddWithValue("@point_n1", TextBox1.Text)
                cmd.Parameters.AddWithValue("@point_n2", TextBox2.Text)
                cmd.Parameters.AddWithValue("@point_n3", TextBox3.Text)
                cmd.Parameters.AddWithValue("@point_n4", TextBox4.Text)
                cmd.Parameters.AddWithValue("@point_n5", TextBox5.Text)
                cmd.Parameters.AddWithValue("@point_n6", TextBox6.Text)
                cmd.Parameters.AddWithValue("@point_n7", TextBox7.Text)
                cmd.Parameters.AddWithValue("@point_n8", TextBox8.Text)
                cmd.Parameters.AddWithValue("@point_n9", TextBox9.Text)
                cmd.Parameters.AddWithValue("@average", TextBox12.Text)
                cmd.Parameters.AddWithValue("@unif", TextBox14.Text)
                cmd.Parameters.AddWithValue("@visa", ComboBox1.Text)
                cmd.Parameters.AddWithValue("@commentaire", TextBox13.Text)

                cmd.ExecuteNonQuery()
                cmd.Parameters.Clear()
                TextBox1.Clear()
                TextBox2.Clear()
                TextBox3.Clear()
                TextBox4.Clear()
                TextBox5.Clear()
                TextBox6.Clear()
                TextBox7.Clear()
                TextBox8.Clear()
                TextBox9.Clear()
                TextBox10.Clear()
                TextBox11.Clear()
                TextBox12.Clear()
                TextBox13.Clear()
                TextBox14.Clear()
                TextBox15.Clear()
                TextBox16.Clear()

                ComboBox1.Text = " "
                Button4.BackColor = Color.Gainsboro
                Button5.BackColor = Color.Gainsboro
                Button6.BackColor = Color.Gainsboro
                Button7.BackColor = Color.Gainsboro
                Button8.BackColor = Color.Gainsboro
                Button9.BackColor = Color.Gainsboro
                Button10.BackColor = Color.Gainsboro
                Button11.BackColor = Color.Gainsboro
                Button12.BackColor = Color.Gainsboro
                MessageBox.Show("Ajouté à Base de Donnée")

            End If

        Else
            MessageBox.Show("La connexion n'est pas ouverte!", "erreur")
        End If
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If SQLcon.State = ConnectionState.Open Then
            ListView1.Items.Clear()
            Button3.BackColor = Color.Green
            Dim cmd As New MySqlCommand("SELECT * FROM `nxq_pm_lampe`", SQLcon)
            Using L As MySqlDataReader = cmd.ExecuteReader()

                While L.Read()
                    Dim dates As String = L("Date").ToString()
                    Dim point1 As String = L("Point_N1")
                    Dim point2 As String = L("Point_N2")
                    Dim point3 As String = L("Point_N3")
                    Dim point4 As String = L("Point_N4")
                    Dim point5 As String = L("Point_N5")
                    Dim point6 As String = L("Point_N6")
                    Dim point7 As String = L("Point_N7")
                    Dim point8 As String = L("Point_N8")
                    Dim point9 As String = L("Point_N9")
                    Dim average As String = L("Average")
                    Dim unif As String = L("Unif")
                    Dim visa As String = L("Visa")
                    Dim commentaire As String = L("Commentaire")


                    ListView1.Items.Add(New ListViewItem(New String() {dates, point1, point2, point3, point4, point5, point6, point7, point8, point9, average, unif, visa, commentaire}))
                End While

            End Using
        ElseIf SQLcon.State = ConnectionState.Closed Then
            ListView1.Items.Clear()
            Button3.BackColor = Color.Yellow
            TextBox1.Text = " "
            TextBox2.Text = " "
            TextBox3.Text = " "
            TextBox4.Text = " "
            TextBox5.Text = " "
            TextBox6.Text = " "
            TextBox7.Text = " "
            TextBox8.Text = " "
            TextBox9.Text = " "
            TextBox13.Text = " "
            ComboBox1.Text = " "
            TextBox12.Text = " "
            TextBox14.Text = " "
            TextBox15.Text = " "
            TextBox16.Text = " "
            TextBox10.Text = " "

        End If
    End Sub





    Private Sub Button14_Click(sender As Object, e As EventArgs) Handles Button14.Click
        CultureInfo.CurrentCulture = New CultureInfo("en-GB")
        If TextBox10.Text.Trim() = "" Then
            MessageBox.Show("Entrez une Date faire un click sur Date")
        ElseIf TextBox1.Text.Trim() = "" Then
            MessageBox.Show("Entrez LE POINT N°1")
        ElseIf TextBox2.Text.Trim() = "" Then
            MessageBox.Show("Entrez LE POINT N°2")
        ElseIf TextBox3.Text.Trim() = "" Then
            MessageBox.Show("Entrez LE POINT N°3")
        ElseIf TextBox4.Text.Trim() = "" Then
            MessageBox.Show("Entrez LE POINT N°4")
        ElseIf TextBox5.Text.Trim() = "" Then
            MessageBox.Show("Entrez LE POINT N°5")
        ElseIf TextBox6.Text.Trim() = "" Then
            MessageBox.Show("Entrez LE POINT N°6")
        ElseIf TextBox7.Text.Trim() = "" Then
            MessageBox.Show("Entrez LE POINT N°7")
        ElseIf TextBox8.Text.Trim() = "" Then
            MessageBox.Show("Entrez LE POINT N°8")
        ElseIf TextBox9.Text.Trim() = "" Then
            MessageBox.Show("Entrez LE POINT N°9")
        End If

        Dim n1 As Double = TextBox1.Text
        Dim n2 As Double = TextBox2.Text
        Dim n3 As Double = TextBox3.Text
        Dim n4 As Double = TextBox4.Text
        Dim n5 As Double = TextBox5.Text
        Dim n6 As Double = TextBox6.Text
        Dim n7 As Double = TextBox7.Text
        Dim n8 As Double = TextBox8.Text
        Dim n9 As Double = TextBox9.Text

        Dim resultat As Double = New Double()

        resultat = (n1 + n2 + n3 + n4 + n5 + n6 + n7 + n8 + n9) / 9
        TextBox12.Text = resultat.ToString("N2")

        If n1 >= n2 And n1 >= n3 And n1 >= n4 And n1 >= n5 And n1 >= n6 And n1 >= n7 And n1 >= n8 And n1 >= n9 Then
            TextBox16.Text = n1
        ElseIf n2 >= n1 And n2 >= n3 And n2 >= n4 And n2 >= n5 And n2 >= n6 And n2 >= n7 And n2 >= n8 And n2 >= n9 Then
            TextBox16.Text = n2
        ElseIf n3 >= n1 And n3 >= n2 And n3 >= n4 And n3 >= n5 And n3 >= n6 And n3 >= n7 And n3 >= n8 And n3 >= n9 Then
            TextBox16.Text = n3
        ElseIf n4 >= n1 And n4 >= n3 And n4 >= n3 And n4 >= n5 And n4 >= n6 And n4 >= n7 And n4 >= n8 And n4 >= n9 Then
            TextBox16.Text = n4
        ElseIf n5 >= n1 And n5 >= n3 And n5 >= n4 And n5 >= n2 And n5 >= n6 And n5 >= n7 And n5 >= n8 And n5 >= n9 Then
            TextBox16.Text = n5
        ElseIf n6 >= n1 And n6 >= n3 And n6 >= n4 And n6 >= n5 And n6 >= n2 And n6 >= n7 And n6 >= n8 And n6 >= n9 Then
            TextBox16.Text = n6
        ElseIf n7 >= n1 And n7 >= n3 And n7 >= n4 And n7 >= n5 And n7 >= n6 And n7 >= n2 And n7 >= n8 And n7 >= n9 Then
            TextBox16.Text = n7
        ElseIf n8 >= n1 And n8 >= n3 And n8 >= n4 And n8 >= n5 And n8 >= n6 And n8 >= n7 And n8 >= n2 And n8 >= n9 Then
            TextBox16.Text = n8
        ElseIf n9 >= n1 And n9 >= n3 And n9 >= n4 And n9 >= n5 And n9 >= n6 And n9 >= n7 And n9 >= n8 And n9 >= n2 Then
            TextBox16.Text = n9


        End If

        If n1 <= n2 And n1 <= n3 And n1 <= n4 And n1 <= n5 And n1 <= n6 And n1 <= n7 And n1 <= n8 And n1 <= n9 Then
            TextBox15.Text = n1
        ElseIf n2 <= n1 And n2 <= n3 And n2 <= n4 And n2 <= n5 And n2 <= n6 And n2 <= n7 And n2 <= n8 And n2 <= n9 Then
            TextBox15.Text = n2
        ElseIf n3 <= n1 And n3 <= n2 And n3 <= n4 And n3 <= n5 And n3 <= n6 And n3 <= n7 And n3 <= n8 And n3 <= n9 Then
            TextBox15.Text = n3
        ElseIf n4 <= n1 And n4 <= n3 And n4 <= n3 And n4 <= n5 And n4 <= n6 And n4 <= n7 And n4 <= n8 And n4 <= n9 Then
            TextBox15.Text = n4
        ElseIf n5 <= n1 And n5 <= n3 And n5 <= n4 And n5 <= n2 And n5 <= n6 And n5 <= n7 And n5 <= n8 And n5 <= n9 Then
            TextBox15.Text = n5
        ElseIf n6 <= n1 And n6 <= n3 And n6 <= n4 And n6 <= n5 And n6 <= n2 And n6 <= n7 And n6 <= n8 And n6 <= n9 Then
            TextBox15.Text = n6
        ElseIf n7 <= n1 And n7 <= n3 And n7 <= n4 And n7 <= n5 And n7 <= n6 And n7 <= n2 And n7 <= n8 And n7 <= n9 Then
            TextBox15.Text = n7
        ElseIf n8 <= n1 And n8 <= n3 And n8 <= n4 And n8 <= n5 And n8 <= n6 And n8 <= n7 And n8 <= n2 And n8 <= n9 Then
            TextBox15.Text = n8
        ElseIf n9 <= n1 And n9 <= n3 And n9 <= n4 And n9 <= n5 And n9 <= n6 And n9 <= n7 And n9 <= n8 And n9 <= n2 Then
            TextBox15.Text = n9
        End If

        resultat = ((TextBox16.Text - TextBox15.Text) / TextBox12.Text) * 50

        TextBox14.Text = resultat.ToString("N2")

        If TextBox14.Text > 5.0 Then
            MessageBox.Show("lampe à changer")
            TextBox13.Text = TextBox13.Text + " " + "lampe à changer = " & TextBox14.Text
        End If

    End Sub

    Private Sub Button13_Click(sender As Object, e As EventArgs) Handles Button13.Click


        If SQLcon.State = ConnectionState.Open Then
            ListView1.Items.Clear()
            Button13.BackColor = Color.Green

            Dim cmd As New MySqlCommand("SELECT * FROM `nxq_pm_lampe`where `Date` between' " + DateTimePicker2.Value.ToString("yyyy-MM-dd 00:00:00") + "' AND '" + DateTimePicker3.Value.ToString("yyyy-MM-dd 23:59:59") + "'", SQLcon)

            Using L As MySqlDataReader = cmd.ExecuteReader()

                While L.Read()
                    Dim dates As String = L("Date").ToString()
                    Dim point1 As String = L("Point_N1")
                    Dim point2 As String = L("Point_N2")
                    Dim point3 As String = L("Point_N3")
                    Dim point4 As String = L("Point_N4")
                    Dim point5 As String = L("Point_N5")
                    Dim point6 As String = L("Point_N6")
                    Dim point7 As String = L("Point_N7")
                    Dim point8 As String = L("Point_N8")
                    Dim point9 As String = L("Point_N9")
                    Dim average As String = L("Average")
                    Dim unif As String = L("Unif")
                    Dim visa As String = L("Visa")
                    Dim commentaire As String = L("Commentaire")


                    ListView1.Items.Add(New ListViewItem(New String() {dates, point1, point2, point3, point4, point5, point6, point7, point8, point9, average, unif, visa, commentaire}))



                End While

            End Using



        ElseIf SQLcon.State = ConnectionState.Closed Then
            ListView1.Items.Clear()
            Button13.BackColor = Color.Yellow
        End If
    End Sub



    Private Sub Label14_Click(sender As Object, e As EventArgs) Handles Label14.Click
        TextBox12.Text = " "
    End Sub

    Private Sub Label15_Click(sender As Object, e As EventArgs) Handles Label15.Click
        TextBox14.Text = " "
    End Sub

    Private Sub Label19_Click(sender As Object, e As EventArgs) Handles Label19.Click
        TextBox15.Text = " "
    End Sub

    Private Sub Label20_Click(sender As Object, e As EventArgs) Handles Label20.Click
        TextBox16.Text = " "
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged

        CultureInfo.CurrentCulture = New CultureInfo("en-GB")
        If TextBox1.Text = " " Then

            Button5.BackColor = Color.Gainsboro
        ElseIf TextBox1.Text < 7.0 Then
            Button5.BackColor = Color.Red
        ElseIf TextBox1.Text > 9.0 Then
            Button5.BackColor = Color.Red

        Else Button5.BackColor = Color.Green
        End If


    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged
        CultureInfo.CurrentCulture = New CultureInfo("en-GB")
        If TextBox2.Text = " " Then

            Button4.BackColor = Color.Gainsboro
        ElseIf TextBox2.Text < 7.0 Then
            Button4.BackColor = Color.Red
        ElseIf TextBox2.Text > 9.0 Then
            Button4.BackColor = Color.Red

        Else Button4.BackColor = Color.Green
        End If
    End Sub

    Private Sub TextBox3_TextChanged(sender As Object, e As EventArgs) Handles TextBox3.TextChanged
        CultureInfo.CurrentCulture = New CultureInfo("en-GB")
        If TextBox3.Text = " " Then

            Button6.BackColor = Color.Gainsboro
        ElseIf TextBox3.Text < 7.0 Then
            Button6.BackColor = Color.Red
        ElseIf TextBox3.Text > 9.0 Then
            Button6.BackColor = Color.Red

        Else Button6.BackColor = Color.Green
        End If
    End Sub

    Private Sub TextBox4_TextChanged(sender As Object, e As EventArgs) Handles TextBox4.TextChanged
        CultureInfo.CurrentCulture = New CultureInfo("en-GB")
        If TextBox4.Text = " " Then

            Button7.BackColor = Color.Gainsboro
        ElseIf TextBox4.Text < 7.0 Then
            Button7.BackColor = Color.Red
        ElseIf TextBox4.Text > 9.0 Then
            Button7.BackColor = Color.Red

        Else Button7.BackColor = Color.Green
        End If
    End Sub

    Private Sub TextBox5_TextChanged(sender As Object, e As EventArgs) Handles TextBox5.TextChanged
        CultureInfo.CurrentCulture = New CultureInfo("en-GB")
        If TextBox5.Text = " " Then

            Button8.BackColor = Color.Gainsboro
        ElseIf TextBox5.Text < 7.0 Then
            Button8.BackColor = Color.Red
        ElseIf TextBox5.Text > 9.0 Then
            Button8.BackColor = Color.Red

        Else Button8.BackColor = Color.Green
        End If
    End Sub

    Private Sub TextBox6_TextChanged(sender As Object, e As EventArgs) Handles TextBox6.TextChanged
        CultureInfo.CurrentCulture = New CultureInfo("en-GB")
        If TextBox6.Text = " " Then

            Button9.BackColor = Color.Gainsboro
        ElseIf TextBox6.Text < 7.0 Then
            Button9.BackColor = Color.Red
        ElseIf TextBox6.Text > 9.0 Then
            Button9.BackColor = Color.Red

        Else Button9.BackColor = Color.Green
        End If
    End Sub

    Private Sub TextBox7_TextChanged(sender As Object, e As EventArgs) Handles TextBox7.TextChanged
        CultureInfo.CurrentCulture = New CultureInfo("en-GB")
        If TextBox7.Text = " " Then

            Button10.BackColor = Color.Gainsboro
        ElseIf TextBox7.Text < 7.0 Then
            Button10.BackColor = Color.Red
        ElseIf TextBox7.Text > 9.0 Then
            Button10.BackColor = Color.Red

        Else Button10.BackColor = Color.Green
        End If
    End Sub

    Private Sub TextBox8_TextChanged(sender As Object, e As EventArgs) Handles TextBox8.TextChanged
        CultureInfo.CurrentCulture = New CultureInfo("en-GB")
        If TextBox8.Text = " " Then

            Button11.BackColor = Color.Gainsboro
        ElseIf TextBox8.Text < 7.0 Then
            Button11.BackColor = Color.Red
        ElseIf TextBox8.Text > 9.0 Then
            Button11.BackColor = Color.Red

        Else Button11.BackColor = Color.Green
        End If
    End Sub

    Private Sub TextBox9_TextChanged(sender As Object, e As EventArgs) Handles TextBox9.TextChanged
        CultureInfo.CurrentCulture = New CultureInfo("en-GB")
        If TextBox9.Text = " " Then

            Button12.BackColor = Color.Gainsboro
        ElseIf TextBox9.Text < 7.0 Then
            Button12.BackColor = Color.Red
        ElseIf TextBox9.Text > 9.0 Then
            Button12.BackColor = Color.Red

        Else Button12.BackColor = Color.Green
        End If
    End Sub
    Private Sub ListView1_SelectedIndexClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ListView1.SelectedIndexChanged
        CultureInfo.CurrentCulture = New CultureInfo("en-GB")
        If ListView1.SelectedItems.Count.ToString("N2") Then
            TextBox10.Text = ListView1.SelectedItems(0).SubItems(0).Text
            TextBox1.Text = ListView1.SelectedItems(0).SubItems(1).Text
            TextBox2.Text = ListView1.SelectedItems(0).SubItems(2).Text
            TextBox3.Text = ListView1.SelectedItems(0).SubItems(3).Text
            TextBox4.Text = ListView1.SelectedItems(0).SubItems(4).Text
            TextBox5.Text = ListView1.SelectedItems(0).SubItems(5).Text
            TextBox6.Text = ListView1.SelectedItems(0).SubItems(6).Text
            TextBox7.Text = ListView1.SelectedItems(0).SubItems(7).Text
            TextBox8.Text = ListView1.SelectedItems(0).SubItems(8).Text
            TextBox9.Text = ListView1.SelectedItems(0).SubItems(9).Text
            ComboBox1.Text = ListView1.SelectedItems(0).SubItems(12).Text
            TextBox13.Text = ListView1.SelectedItems(0).SubItems(13).Text
        End If

        TextBox1.Text = Replace(TextBox1.Text, ",", ".")
        TextBox2.Text = Replace(TextBox2.Text, ",", ".")
        TextBox3.Text = Replace(TextBox3.Text, ",", ".")
        TextBox4.Text = Replace(TextBox4.Text, ",", ".")
        TextBox5.Text = Replace(TextBox5.Text, ",", ".")
        TextBox6.Text = Replace(TextBox6.Text, ",", ".")
        TextBox7.Text = Replace(TextBox7.Text, ",", ".")
        TextBox8.Text = Replace(TextBox8.Text, ",", ".")
        TextBox9.Text = Replace(TextBox9.Text, ",", ".")
        TextBox12.Text = " "
        TextBox14.Text = " "
        TextBox15.Text = " "
        TextBox16.Text = " "

    End Sub


End Class


