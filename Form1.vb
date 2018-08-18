Imports System.IO

Public Class Form1
    Public mytracks As New List(Of EDLTrack)
    Public defaudio As String = "aac"
    Public mydroplist As New List(Of String)
    Public allowmuldrop As Boolean

    Private Sub Button2_Click(sender As Object, e As EventArgs)
        Dim mydigšit As New OpenFileDialog
        mydigšit.Filter = "*.edf|EDF file (*.edf)|*.*|All files (*.*)"
        mydigšit.ShowDialog()
        If Not mydigšit.FileName = "" Then
            TextBox1.Text = mydigšit.FileName
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs)
        Dim mydigšit As New SaveFileDialog
        mydigšit.ShowDialog()
        If Not mydigšit.FileName = "" Then
            TextBox2.Text = mydigšit.FileName
        End If
    End Sub

    Private Sub Form1_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        RecheckOptions()
        Dim mywriter As New StreamWriter(Application.StartupPath + "/options")
        mywriter.WriteLine(TextBox1.Text)
        mywriter.WriteLine(TextBox2.Text)
        mywriter.WriteLine(TextBox3.Text)
        mywriter.WriteLine(TextBox5.Text)
        mywriter.WriteLine(TextBox6.Text)
        mywriter.Close()
    End Sub

    Private Sub Form1_Shown(sender As Object, e As EventArgs) Handles Me.Shown
        If File.Exists("options") Then
            Dim myreader As New StreamReader(Application.StartupPath + "/options")
            TextBox1.Text = myreader.ReadLine()
            TextBox2.Text = myreader.ReadLine()
            TextBox3.Text = myreader.ReadLine()
            TextBox5.Text = myreader.ReadLine()
            TextBox6.Text = myreader.ReadLine()
            myreader.Close()
        End If
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs)
        Dim mydigšit As New OpenFileDialog
        mydigšit.Filter = "*.exe|Executable file (*.exe)|*.*|All files (*.*)"
        mydigšit.ShowDialog()
        If Not mydigšit.FileName = "" Then
            TextBox3.Text = mydigšit.FileName
        End If
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim mypic As Bitmap
        mypic = New Bitmap(480, 480)
        Dim mygraph As Graphics
        mygraph = Graphics.FromImage(mypic)
        mygraph.FillRectangle(New SolidBrush(Color.Green), New Rectangle(0, 0, mypic.Width, mypic.Height))
        mygraph.FillEllipse(New SolidBrush(Color.Red), New Rectangle(0, 0, mypic.Width, mypic.Height))
        'mygraph.FillEllipse(New SolidBrush(Color.Yellow), New Rectangle(mypic.Width / 4, mypic.Width / 4, mypic.Width / 2, mypic.Height / 2))
        mygraph.DrawEllipse(New Pen(New SolidBrush(Color.Black), 3), New Rectangle(0, 0, mypic.Width, mypic.Height))
        PictureBox1.Image = mypic
        Dim dragdroptxt As String() = New String() {"TextBox1", "TextBox2", "TextBox3", "TextBox6"}
        Dim allcontrols As List(Of Control)
        allcontrols = GetControls(Me)
        For Each item In allcontrols
            If TypeOf item Is TextBox Then
                If DirectCast(item, TextBox).Multiline Then
                    AddHandler DirectCast(item, TextBox).KeyDown, AddressOf slttext
                End If
                If dragdroptxt.Where(Function(x) x.ToLower() = item.Name.ToLower()).Count > 0 Then
                    AddHandler item.DragEnter, AddressOf TextBoxD_DragEnter
                    AddHandler item.DragDrop, AddressOf TextBoxD_DragDrop
                End If
            End If
        Next
    End Sub

    Public Function GetControls(myctl As Object) As List(Of Control)
        Dim mylist As List(Of Control)
        mylist = New List(Of Control)
        For Each mycontrol In myctl.Controls
            mylist.AddRange(GetControls(mycontrol))
            mylist.Add(mycontrol)
        Next
        Return mylist
    End Function

    Private Sub slttext(sender As Object, e As KeyEventArgs)
        If e.Control AndAlso e.KeyCode = Keys.A Then
            Dim txt As TextBox
            txt = DirectCast(sender, TextBox)
            txt.SelectionStart = 0
            txt.SelectionLength = txt.Text.Length
        End If
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        RecheckOptions()
        Button4.Enabled = False
        Button4.Text = "Synchronizing..."
        Dim myedlreader As New StreamReader(TextBox1.Text)
        Dim myformat As String
        myformat = myedlreader.ReadLine()
        mytracks.Clear()
        While Not myedlreader.EndOfStream
            Dim myline As String
            myline = myedlreader.ReadLine()
            If Not myline = "" AndAlso Not myline = Environment.NewLine Then
                mytracks.Add(New EDLTrack(myformat, myline))
            End If
        End While
        For Each item In mytracks
            'If Integer.Parse(item.GetOption("Track")) > 1 Then
            'End If
            Dim myff As New ProcessStartInfo
            myff.FileName = TextBox3.Text
            If item.GetOption("MediaType").ToLower() = "audio" Then
                Dim tm1 As TimeSpan
                Dim tm2 As TimeSpan
                tm1 = TimeSpan.FromMilliseconds(TrimTo(item.GetOption("StreamStart"), New String() {".", ","}))
                tm2 = TimeSpan.FromMilliseconds(TrimTo(item.GetOption("StreamLength"), New String() {".", ","}))
                Dim mynumb As Double
                mynumb = Double.Parse(TrimTo(item.GetOption("StreamLength"), New String() {".", ","})) / Double.Parse(TrimTo(item.GetOption("Length"), New String() {".", ","}))
                myff.Arguments = IfAdd(tm1.TotalSeconds > Double.Parse(TextBox5.Text), "-ss " + FormatMeDate(TimeSpan.FromMilliseconds(tm1.TotalMilliseconds - Double.Parse(TextBox5.Text) * 1000), "ms") + " ") + "-i " + ChrW(34) + item.GetOption("FileName") + ChrW(34) + IfAdd(tm1.TotalMilliseconds > 0, IfElse(tm1.TotalMilliseconds > Double.Parse(TextBox5.Text) * 1000, " -ss " + FormatMeDate(TimeSpan.FromSeconds(Double.Parse(TextBox5.Text)), "ms"), " -ss " + FormatMeDate(tm1, "ms"))) + " -t " + FormatMeDate(tm2, "ms") + IfElse(Not mynumb = 1, " -filter:a " + ChrW(34) + MaximizeTo(mynumb, 0.5, 2, ",", "atempo="), " -c:a copy") + ChrW(34) + " -vn " + ChrW(34) + TextBox6.Text + "/" + item.GetOption("ID") + "." + defaudio + ChrW(34)
                If File.Exists(TextBox6.Text + "/" + item.GetOption("ID") + "." + defaudio) Then
                    File.Delete(TextBox6.Text + "/" + item.GetOption("ID") + "." + defaudio)
                End If
            Else
                Dim tm1 As TimeSpan
                Dim tm2 As TimeSpan
                tm1 = TimeSpan.FromMilliseconds(TrimTo(item.GetOption("StreamStart"), New String() {".", ","}))
                tm2 = TimeSpan.FromMilliseconds(TrimTo(item.GetOption("StreamLength"), New String() {".", ","}))
                Dim mynumb As Double
                mynumb = Double.Parse(TrimTo(item.GetOption("Length"), New String() {".", ","})) / Double.Parse(TrimTo(item.GetOption("StreamLength"), New String() {".", ","}))
                myff.Arguments = IfAdd(tm1.TotalSeconds > Double.Parse(TextBox5.Text), "-ss " + FormatMeDate(TimeSpan.FromMilliseconds(tm1.TotalMilliseconds - Double.Parse(TextBox5.Text) * 1000), "ms") + " ") + "-i " + ChrW(34) + item.GetOption("FileName") + ChrW(34) + IfAdd(tm1.TotalMilliseconds > 0, IfElse(tm1.TotalMilliseconds > Double.Parse(TextBox5.Text) * 1000, " -ss " + FormatMeDate(TimeSpan.FromSeconds(Double.Parse(TextBox5.Text)), "ms"), " -ss " + FormatMeDate(tm1, "ms"))) + " -t " + FormatMeDate(tm2, "ms") + IfElse(Not mynumb = 1, " -filter:v " + ChrW(34) + "setpts=" + mynumb.ToString().Replace(",", ".") + "*PTS" + ChrW(34), " -c:v copy") + " -an " + ChrW(34) + TextBox6.Text + "/" + item.GetOption("ID") + "." + Path.GetExtension(item.GetOption("FileName")).Replace(".", "").ToLower() + ChrW(34)
                If File.Exists(TextBox6.Text + "/" + item.GetOption("ID") + "." + Path.GetExtension(item.GetOption("FileName")).Replace(".", "").ToLower()) Then
                    File.Delete(TextBox6.Text + "/" + item.GetOption("ID") + "." + Path.GetExtension(item.GetOption("FileName")).Replace(".", "").ToLower())
                End If
            End If
            Dim myprocess As Process
            myprocess = Process.Start(myff)
            While Not myprocess.HasExited
                Application.DoEvents()
            End While
        Next
        mytracks = mytracks.OrderBy(Function(x) TrimTo(x.GetOption("StreamStart"), New String() {".", ","})).ToList()
        Dim videotrackcount As Integer
        Dim audiotrackcount As Integer
        videotrackcount = 0
        audiotrackcount = 0
        Dim lastaudioitem As EDLTrack
        Dim lastvideoitem As EDLTrack
        Dim myauwriter As New StreamWriter(TextBox6.Text + "/concau")
        Dim myviwriter As New StreamWriter(TextBox6.Text + "/concvi")
        For Each mnitem In mytracks
            If mnitem.GetOption("MediaType").ToLower() = "audio" Then
                myauwriter.WriteLine("file '" + mnitem.GetOption("ID") + "." + defaudio + "'")
                lastaudioitem = mnitem
                audiotrackcount += 1
            Else
                myviwriter.WriteLine("file '" + mnitem.GetOption("ID") + "." + Path.GetExtension(mnitem.GetOption("FileName")).Replace(".", "").ToLower() + "'")
                lastvideoitem = mnitem
                videotrackcount += 1
            End If
        Next
        myauwriter.Close()
        myviwriter.Close()
        Dim myff2 As New ProcessStartInfo
        myff2.FileName = TextBox3.Text
        If File.Exists(TextBox6.Text + "/final." + defaudio) Then
            File.Delete(TextBox6.Text + "/final." + defaudio)
        End If
        Dim mylastex As String
        mylastex = Path.GetExtension(lastvideoitem.GetOption("FileName")).Replace(".", "").ToLower()
        If File.Exists(TextBox6.Text + "/final." + mylastex) Then
            File.Delete(TextBox6.Text + "/final." + mylastex)
        End If
        Dim myproc As Process
        If audiotrackcount > 1 Then
            myff2.Arguments = "-f concat -i " + ChrW(34) + TextBox6.Text + "/concau" + ChrW(34) + " -c copy " + ChrW(34) + TextBox6.Text + "/final." + defaudio + ChrW(34)
            myproc = Process.Start(myff2)
            While Not myproc.HasExited
                Application.DoEvents()
            End While
        Else
            File.Move(TextBox6.Text + "/" + lastaudioitem.GetOption("ID") + "." + defaudio, TextBox6.Text + "/final." + defaudio)
        End If
        If videotrackcount > 1 Then
            myff2.Arguments = "-f concat -i " + ChrW(34) + TextBox6.Text + "/concvi" + ChrW(34) + " -c copy " + ChrW(34) + TextBox6.Text + "/final." + mylastex + ChrW(34)
            myproc = Process.Start(myff2)
            While Not myproc.HasExited
                Application.DoEvents()
            End While
        Else
            File.Move(TextBox6.Text + "/" + lastvideoitem.GetOption("ID") + "." + mylastex, TextBox6.Text + "/final." + mylastex)
        End If
        myff2.Arguments = "-i " + ChrW(34) + TextBox6.Text + "/final." + mylastex + ChrW(34) + " -i " + ChrW(34) + TextBox6.Text + "/final." + defaudio + ChrW(34) + " -c copy " + ChrW(34) + TextBox2.Text + ChrW(34)
        myproc = Process.Start(myff2)
        While Not myproc.HasExited
            Application.DoEvents()
        End While
        myedlreader.Close()
        Button4.Text = "Synchronize"
        Button4.Enabled = True
    End Sub

    Private Function IfElse(p As Boolean, v1 As String, v2 As String) As String
        If p Then
            Return v1
        Else
            Return v2
        End If
    End Function

    Private Function IfAdd(p As Boolean, v As String) As String
        Return IfElse(p, v, "")
    End Function

    Private Function MaximizeTo(value As Double, minfactor As Double, maxfactor As Double, splitstr As String, addbefore As String) As String
        Dim curvalue As Double
        curvalue = 1
        Dim mytext As String
        mytext = ""
        If value < minfactor Then
            While curvalue > value
                If curvalue * minfactor >= value Then
                    curvalue = curvalue * minfactor
                    mytext += addbefore + minfactor.ToString().Replace(",", ".") + splitstr
                Else
                    curvalue = value
                    mytext += addbefore + (value / minfactor).ToString().Replace(",", ".") + splitstr
                End If
            End While
        Else
            If value > maxfactor Then
                While curvalue < value
                    If curvalue * maxfactor <= value Then
                        curvalue = curvalue * maxfactor
                        mytext += addbefore + maxfactor.ToString().Replace(",", ".") + splitstr
                    Else
                        curvalue = value
                        mytext += addbefore + (value / maxfactor).ToString().Replace(",", ".") + splitstr
                    End If
                End While
            Else
                mytext = addbefore + value.ToString().Replace(",", ".") + splitstr
            End If
        End If
        Return mytext.Substring(0, mytext.Length - 1)
    End Function

    Private Function TrimTo(v1 As String, v2() As String) As Double
        Dim mycnc As Integer
        mycnc = v1.Length
        For Each v31 In v2
            If v1.Contains(v31) Then
                If v1.IndexOf(v31) < mycnc Then
                    mycnc = v1.IndexOf(v31)
                End If
            End If
        Next
        Return v1.Substring(0, mycnc)
    End Function

    Private Function FormatMeDate(tm1 As TimeSpan, v As String) As String
        If v = "ms" Then
            Return Math.Floor(tm1.TotalHours).ToString().PadLeft(2, "0") + ":" + tm1.Minutes.ToString().PadLeft(2, "0") + ":" + tm1.Seconds.ToString().PadLeft(2, "0") + "." + tm1.Milliseconds.ToString().PadLeft(2, "0")
        End If
        If v = "s" Then
            Return Math.Floor(tm1.TotalHours).ToString().PadLeft(2, "0") + ":" + tm1.Minutes.ToString().PadLeft(2, "0") + ":" + tm1.Seconds.ToString().PadLeft(2, "0")
        End If
        Return tm1.ToString("hh:mm:ss")
    End Function

    Private Sub RecheckOptions()
        TextBox1.Text = TextBox1.Text.Replace("\", "/")
        TextBox2.Text = TextBox2.Text.Replace("\", "/")
        TextBox3.Text = TextBox3.Text.Replace("\", "/")
        TextBox6.Text = TextBox6.Text.Replace("\", "/")
        While TextBox6.Text.EndsWith("/")
            TextBox6.Text = TextBox6.Text.Substring(0, TextBox6.Text.Length)
        End While
        TextBox6.SelectionStart = TextBox6.Text.Length
        TextBox6.SelectionLength = 0
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Dim mydigšit As New FolderBrowserDialog
        mydigšit.ShowDialog()
        If Not mydigšit.SelectedPath = "" Then
            TextBox3.Text = mydigšit.SelectedPath
        End If
    End Sub

    Private Sub TextBoxD_DragEnter(ByVal sender As Object, ByVal e As DragEventArgs)
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            e.Effect = DragDropEffects.All
        End If
    End Sub

    Private Sub TextBoxD_DragDrop(ByVal sender As Object, ByVal e As DragEventArgs)
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            DirectCast(sender, TextBox).Text = e.Data.GetData(DataFormats.FileDrop)(0)
        End If
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Button6.Enabled = False
        PictureBox1.AllowDrop = True
        mydroplist.Clear()
        Label7.Text = "Drag & drop video file"
        While mydroplist.Count = 0
            Application.DoEvents()
        End While
        Label7.Text = "Drag & drop audio file"
        While mydroplist.Count = 1
            Application.DoEvents()
        End While
        Label7.Text = ""
        Dim myff3 As New ProcessStartInfo
        Dim myproc As Process
        Dim i As Integer
        i = 1
        Dim myext As String
        myext = Path.GetExtension(mydroplist(0)).Replace(".", "").ToLower()
        While File.Exists(TextBox6.Text + "/" + i.ToString() + "." + myext)
            i += 1
        End While
        myff3.FileName = TextBox3.Text
        myff3.Arguments = "-i " + ChrW(34) + mydroplist(0) + ChrW(34) + " -an -c:v copy " + ChrW(34) + TextBox6.Text + "/" + i.ToString() + "." + myext + ChrW(34)
        myproc = Process.Start(myff3)
        While Not myproc.HasExited
            Application.DoEvents()
        End While
        myff3.Arguments = "-i " + ChrW(34) + TextBox6.Text + "/" + i.ToString() + "." + myext + ChrW(34) + " -i " + ChrW(34) + mydroplist(1) + ChrW(34) + " -c copy " + ChrW(34) + TextBox6.Text + "/" + (i + 1).ToString() + "." + myext + ChrW(34)
        myproc = Process.Start(myff3)
        While Not myproc.HasExited
            Application.DoEvents()
        End While
        PictureBox1.AllowDrop = False
        Button6.Enabled = True
    End Sub

    Private Sub TextBoxEF_DragEnter(ByVal sender As Object, ByVal e As DragEventArgs) Handles PictureBox1.DragEnter
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            e.Effect = DragDropEffects.All
        End If
    End Sub

    Private Sub TextBoxEF_DragDrop(ByVal sender As Object, ByVal e As DragEventArgs) Handles PictureBox1.DragDrop
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            If Not allowmuldrop Then
                mydroplist.Add(e.Data.GetData(DataFormats.FileDrop)(0))
            Else
                mydroplist.AddRange(e.Data.GetData(DataFormats.FileDrop))
            End If
        End If
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        Button6.Enabled = False
        PictureBox1.AllowDrop = True
        allowmuldrop = True
        mydroplist.Clear()
        Label7.Text = "Drag & drop files to merge"
        While mydroplist.Count = 0
            Application.DoEvents()
        End While
        Label7.Text = ""
        mydroplist = mydroplist.OrderBy(Function(x) Path.GetFileName(x)).tolist()
        Dim myff3 As New ProcessStartInfo
        Dim myproc As Process
        Dim i As Integer
        i = 1
        While File.Exists(TextBox6.Text + "/cn" + i.ToString())
            i += 1
        End While
        Dim i2 As Integer
        i2 = 1
        Dim myext As String
        myext = Path.GetExtension(mydroplist(0)).Replace(".", "").ToLower()
        While File.Exists(TextBox6.Text + "/" + i2.ToString() + "." + myext)
            i2 += 1
        End While
        myff3.FileName = TextBox3.Text
        Dim mymergwriter As New StreamWriter(TextBox6.Text + "/cn" + i.ToString())
        For Each flitem In mydroplist
            mymergwriter.WriteLine("file '" + flitem.Replace("\", "/") + "'")
        Next
        mymergwriter.Close()
        myff3.Arguments = "-f concat -safe 0 -i " + ChrW(34) + (TextBox6.Text + "/cn").Replace("/", "\") + i.ToString() + ChrW(34) + " -c copy " + ChrW(34) + TextBox6.Text + "/" + i2.ToString() + "." + myext + ChrW(34)
        Clipboard.SetText(myff3.Arguments)
        myproc = Process.Start(myff3)
        While Not myproc.HasExited
            Application.DoEvents()
        End While
        allowmuldrop = False
        PictureBox1.AllowDrop = False
        Button6.Enabled = True
    End Sub
End Class
