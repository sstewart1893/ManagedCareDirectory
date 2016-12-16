Imports System.Xml
Imports System.Xml.Xsl
Imports java.io
Imports org.xml.sax
Imports System.Data.SqlClient
Imports System.Net.Mail
Imports GlobalLibrary
Module Module1
    Private DBName As String = "DataRepository"
    'Private DBName As String = "Testing"
    Private CN As String = "Server=NASPROSQL1;Database=" & DBName & ";User Id=HLIDBAdmin;Password=^HLI<&dm!n;Connection Timeout=150"
    Private filePath As String = "\\NASPROSQL1\HLI_Applications\Dashboard\"
    Private filePathTemplates As String = "\\NASPROSQL1\HLI_Applications\Dashboard\Templates\"

    Private GlobalFunctions As GlobalLibrary.Functions

    Private SQLHelper As New GlobalLibrary.SqlHelper
    Private MyDate As Integer

    Sub Main()
        Dim strParameters() As String = Split(Command(), "|")
        If strParameters.Length < 3 Then
            MyDate = 9679
            'MsgBox("You have entered outside of the Dashboard or there was an error loading")
            'Me.Dispose()
            'Exit Sub
        Else
            MyDate = CInt(strParameters(1))
        End If

        GeneratePDFDocument()

    End Sub





    Private Function GeneratePDFDocument() As Boolean

        Try
            Dim ds As DataSet = Nothing
            ds = New DataSet("ExchangeProfile")
            ds = SQLHelper.ExecuteDataset(CN, "RPT.s_GetManagedCareDirectoryHeadersForPDF", MyDate)
            If ds.Tables(0).Rows.Count = 0 Then
                Throw New Exception("No PDF Document Info Found")
            End If

            ds.Relations.Add("Companies", ds.Tables(1).Columns("TableName"), ds.Tables(2).Columns("TableName")).Nested = True
            ds.Relations.Add("States", ds.Tables(3).Columns("TableName"), ds.Tables(4).Columns("TableName")).Nested = True
            ds.Relations.Add("StatesCompanies", ds.Tables(4).Columns("statename"), ds.Tables(5).Columns("statename")).Nested = True

            Dim sXMLFileName As String = "ManagedCareDirectory.xml"
            Dim sXSLFileName As String = "ManagedCarePDF.xsl"
            Dim sFOFileName As String = "ManagedCarePDF.fo"

            Dim xmlDoc As New XmlDataDocument(ds)
            xmlDoc.Save(filePath & sXMLFileName)

            Dim transform As New XslCompiledTransform()
            transform.Load(filePath & "Templates\" & sXSLFileName)
            transform.Transform(filePath & sXMLFileName, filePath & sFOFileName)
            GeneratePDF(filePath & sFOFileName, filePath & "ManagedCareDirectory2.pdf", "GeneratePDF" + MyDate.ToString)
            Return True
        Catch ex As Exception
            SendErrormail(ex.Message, "ManagedCareDirectory Document Generation- GeneratePDFDocument", "")
            Return False

        End Try

    End Function



    ''' <summary>
    ''' Create a .pdf file from the given fo file, provide a meaningful message if there is a problem
    ''' </summary>
    ''' <param name="foFile"></param>
    ''' <param name="pdfFile"></param>
    ''' <param name="strMessage"></param>
    ''' <remarks></remarks>
    Private Sub GeneratePDF(ByVal foFile As String, ByVal pdfFile As String, ByVal strMessage As String)
        Try
            Dim streamFO As New FileInputStream(foFile)

            Dim src As New InputSource(streamFO)
            Dim streamOut As New FileOutputStream(pdfFile)
            Dim driver As New org.apache.fop.apps.Driver(src, streamOut)
            driver.setRenderer(1)
            driver.run()

            streamOut.close()

        Catch ex As Exception
            SendErrormail(ex.Message, strMessage, "ETL")

        End Try


    End Sub


    ''' <summary>
    ''' Send an Error Email to the DBAs
    ''' </summary>
    ''' <param name="myMessage"></param>
    ''' <param name="sSubject"></param>
    ''' <param name="sUser"></param>
    ''' <remarks></remarks>
    Public Sub SendErrormail(ByVal myMessage As String, ByVal sSubject As String, ByVal sUser As String)
        MsgBox(myMessage) 'For testing
        Exit Sub 'For testing

        Dim Message As New MailMessage()

        Try
            Message.IsBodyHtml = True

            Message.From = New System.Net.Mail.MailAddress("DRGSN@Dresources.Com")
            Message.[To].Add(New System.Net.Mail.MailAddress("SStewart@HL-ISY.com", "sstewart"))
            Message.[To].Add(New System.Net.Mail.MailAddress("MSchellhammer@HL-ISY.com", "MSchellhammer"))
            Message.Subject = sSubject & "-" & sUser
            Message.Body = myMessage
            Dim Client As New System.Net.Mail.SmtpClient()
            Using dr As SqlDataReader = GlobalLibrary.SqlHelper.ExecuteReader(CN, "s_Getapplicationvariableinfolist", 0)
                While dr.Read()
                    If dr("VariableName").ToString = "SMTPServer1" Then
                        Client.Host = dr("VariableValue").ToString
                    End If
                End While
            End Using
            Client.UseDefaultCredentials = True

            Client.Send(Message)
        Catch ex As Exception
            'AddToLog("SendErrormail: " & ex.ToString)
        End Try

    End Sub
End Module
