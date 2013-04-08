Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    DefaultView =0
    TabularFamily =0
    DateGrouping =1
    GrpKeepTogether =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =9360
    DatasheetFontHeight =10
    ItemSuffix =11
    Left =210
    Top =150
    Right =11370
    Bottom =5355
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x0f53975bad60e240
    End
    RecordSource ="people (sample table)"
    Caption ="people"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0xa0050000a0050000a0050000a005000000000000902400006801000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    FilterOnLoad =0
    AllowLayoutView =0
    DatasheetGridlinesColor12 =12632256
    Begin
        Begin Label
            FontItalic = NotDefault
            BackStyle =0
            TextAlign =1
            TextFontFamily =18
            FontSize =11
            FontWeight =700
            ForeColor =8388608
            FontName ="Times New Roman"
        End
        Begin Rectangle
            BackStyle =0
            BorderWidth =1
            BorderLineStyle =0
            BorderColor =8388608
        End
        Begin Line
            BorderLineStyle =0
            BorderColor =8388608
        End
        Begin Image
            OldBorderStyle =0
            BorderLineStyle =0
            PictureAlignment =2
        End
        Begin CheckBox
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
        End
        Begin TextBox
            OldBorderStyle =0
            BorderLineStyle =0
            BackStyle =0
            FontName ="Arial"
        End
        Begin ListBox
            OldBorderStyle =0
            BorderLineStyle =0
            FontName ="Arial"
        End
        Begin ComboBox
            OldBorderStyle =0
            BorderLineStyle =0
            BackStyle =0
            FontName ="Arial"
        End
        Begin Subform
            OldBorderStyle =0
            BorderLineStyle =0
        End
        Begin BreakLevel
            ControlSource ="Name"
        End
        Begin BreakLevel
            ControlSource ="Birthday"
        End
        Begin FormHeader
            KeepTogether = NotDefault
            Height =924
            Name ="ReportHeader"
            Begin
                Begin Label
                    BackStyle =1
                    Left =60
                    Top =60
                    Width =1140
                    Height =504
                    FontSize =20
                    Name ="Label6"
                    Caption ="people"
                End
            End
        End
        Begin PageHeader
            Height =408
            Name ="PageHeaderSection"
            Begin
                Begin Label
                    Left =60
                    Top =60
                    Width =2208
                    Height =288
                    Name ="Name_Label"
                    Caption ="Name"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =3
                    Left =2328
                    Top =60
                    Width =984
                    Height =288
                    Name ="Birthday_Label"
                    Caption ="Birthday"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    Left =3372
                    Top =60
                    Width =4368
                    Height =288
                    Name ="Comment_Label"
                    Caption ="Comment"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    BorderWidth =2
                    Left =60
                    Top =348
                    Width =7680
                    Name ="Line9"
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            Height =360
            Name ="Detail"
            Begin
                Begin TextBox
                    Left =60
                    Top =60
                    Width =2208
                    Name ="Name"
                    ControlSource ="Name"

                End
                Begin TextBox
                    Left =2328
                    Top =60
                    Width =984
                    TabIndex =1
                    Name ="Birthday"
                    ControlSource ="Birthday"
                    Format ="Short Date"

                End
                Begin TextBox
                    Left =3372
                    Top =60
                    Width =4368
                    TabIndex =2
                    Name ="Comment"
                    ControlSource ="Comment"

                End
            End
        End
        Begin PageFooter
            Height =516
            Name ="PageFooterSection"
            Begin
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =1
                    TextFontFamily =18
                    Left =60
                    Top =240
                    Width =4560
                    Height =276
                    FontSize =9
                    FontWeight =700
                    ForeColor =8388608
                    Name ="Text7"
                    ControlSource ="=Now()"
                    Format ="Long Date"
                    FontName ="Times New Roman"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =3
                    TextFontFamily =18
                    Left =4740
                    Top =240
                    Width =4560
                    Height =276
                    FontSize =9
                    FontWeight =700
                    TabIndex =1
                    ForeColor =8388608
                    Name ="Text8"
                    ControlSource ="=\"Page \" & [Page] & \" of \" & [Pages]"
                    FontName ="Times New Roman"

                End
                Begin Line
                    BorderWidth =3
                    Left =60
                    Top =240
                    Width =9240
                    BorderColor =12632256
                    Name ="Line10"
                End
            End
        End
        Begin FormFooter
            KeepTogether = NotDefault
            Height =0
            Name ="ReportFooter"
        End
    End
End
CodeBehindForm
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Compare Database

Private StartDate As Variant
Private EndDate As Variant

Private Sub Report_Open(Cancel As Integer)

StartDate = InputDate("Select Start Date")
EndDate = InputDate("Select End Date")
If IsDate(StartDate) And IsDate(EndDate) Then
    Me.Filter = "Birthday Between #" & StartDate & "# And #" & _
        EndDate & "#"
    Me.FilterOn = True
End If
End Sub
