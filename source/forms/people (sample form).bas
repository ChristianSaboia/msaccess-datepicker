Version =20
VersionRequired =20
Begin Form
    AutoCenter = NotDefault
    DefaultView =0
    TabularFamily =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =6996
    DatasheetFontHeight =10
    ItemSuffix =7
    Left =4755
    Top =2625
    Right =12030
    Bottom =4005
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xa4027657ad60e240
    End
    RecordSource ="people (sample table)"
    Caption ="people"
    DatasheetFontName ="Arial"
    FilterOnLoad =0
    AllowLayoutView =0
    DatasheetGridlinesColor12 =12632256
    Begin
        Begin Label
            BackStyle =0
            BackColor =-2147483633
            ForeColor =-2147483630
        End
        Begin Rectangle
            SpecialEffect =3
            BackStyle =0
            BorderLineStyle =0
        End
        Begin Image
            BackStyle =0
            OldBorderStyle =0
            BorderLineStyle =0
            PictureAlignment =2
        End
        Begin CommandButton
            FontSize =8
            FontWeight =400
            FontName ="MS Sans Serif"
            BorderLineStyle =0
        End
        Begin OptionButton
            SpecialEffect =2
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
        End
        Begin CheckBox
            SpecialEffect =2
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
        End
        Begin OptionGroup
            SpecialEffect =3
            BorderLineStyle =0
        End
        Begin BoundObjectFrame
            SpecialEffect =2
            OldBorderStyle =0
            BorderLineStyle =0
            BackStyle =0
        End
        Begin TextBox
            SpecialEffect =2
            BorderLineStyle =0
            BackColor =-2147483643
            ForeColor =-2147483640
        End
        Begin ListBox
            SpecialEffect =2
            BorderLineStyle =0
            BackColor =-2147483643
            ForeColor =-2147483640
        End
        Begin ComboBox
            SpecialEffect =2
            BorderLineStyle =0
            BackColor =-2147483643
            ForeColor =-2147483640
        End
        Begin Subform
            SpecialEffect =2
            BorderLineStyle =0
        End
        Begin UnboundObjectFrame
            SpecialEffect =2
            OldBorderStyle =1
        End
        Begin ToggleButton
            FontSize =8
            FontWeight =400
            FontName ="MS Sans Serif"
            BorderLineStyle =0
        End
        Begin Tab
            BackStyle =0
            BorderLineStyle =0
        End
        Begin FormHeader
            Height =0
            BackColor =-2147483633
            Name ="FormHeader"
        End
        Begin Section
            Height =1392
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    Left =1848
                    Top =120
                    Width =2568
                    ColumnWidth =2568
                    Name ="Name"
                    ControlSource ="Name"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =60
                            Top =120
                            Width =1728
                            Height =240
                            Name ="Name_Label"
                            Caption ="Name"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    Left =1848
                    Top =480
                    Width =1140
                    ColumnWidth =1140
                    TabIndex =1
                    Name ="Birthday"
                    ControlSource ="Birthday"
                    Format ="Short Date"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =60
                            Top =480
                            Width =1728
                            Height =240
                            Name ="Birthday_Label"
                            Caption ="Birthday"
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    Left =1848
                    Top =840
                    Width =5088
                    Height =432
                    ColumnWidth =3000
                    TabIndex =2
                    Name ="Comment"
                    ControlSource ="Comment"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =60
                            Top =840
                            Width =1728
                            Height =240
                            Name ="Comment_Label"
                            Caption ="Comment"
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =3060
                    Top =480
                    Width =246
                    Height =246
                    TabIndex =3
                    Name ="cmdBirthday"
                    OnClick ="[Event Procedure]"
                    PictureData = Begin
                        0x2800000010000000100000000100040000000000800000000000000000000000 ,
                        0x0000000000000000000000000000800000800000008080008000000080008000 ,
                        0x8080000080808000c0c0c0000000ff00c0c0c00000ffff00ff000000c0c0c000 ,
                        0xffff0000ffffff00dadadadadadadada000000000000000d0fffffffffffff0a ,
                        0x0f7777777fffff0d0f7f7f7f7fffff0a0f77777777777f0d0f7f7f7f7f7f7f0a ,
                        0x0f77777777777f0d0f7f7f7f7f7f7f0a0f77777777777f0d0f7f7f7f7f7f7f0a ,
                        0x0f77777777777f0d0fffffffffffff0a0f777777fff77f0d0fffffffffffff0a ,
                        0x000000000000000d000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000
                    End
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
            End
        End
        Begin FormFooter
            Height =0
            BackColor =-2147483633
            Name ="FormFooter"
        End
    End
End
CodeBehindForm
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Compare Database

Private Sub cmdBirthday_Click()
InputDateField Birthday, "Select Birthday"
End Sub
