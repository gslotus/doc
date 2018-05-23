object fmsign: Tfmsign
  Left = 139
  Top = 119
  Width = 696
  Height = 424
  BorderIcons = [biSystemMenu]
  Caption = #36039#26009#31805#26680#20316#26989
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -12
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  Position = poScreenCenter
  OnShow = FormShow
  PixelsPerInch = 106
  TextHeight = 13
  object Button1: TButton
    Left = 0
    Top = 0
    Width = 113
    Height = 25
    Caption = #36039#26009#25918#34892
    Font.Charset = ANSI_CHARSET
    Font.Color = clWindowText
    Font.Height = -16
    Font.Name = #26032#32048#26126#39636
    Font.Style = []
    ParentFont = False
    TabOrder = 0
    OnClick = Button1Click
  end
  object DBGrid1: TDBGrid
    Left = 0
    Top = 25
    Width = 661
    Height = 370
    Align = alBottom
    DataSource = DataSource1
    Font.Charset = ANSI_CHARSET
    Font.Color = clWindowText
    Font.Height = -16
    Font.Name = #26032#32048#26126#39636
    Font.Style = []
    ParentFont = False
    ReadOnly = True
    TabOrder = 1
    TitleFont.Charset = ANSI_CHARSET
    TitleFont.Color = clWindowText
    TitleFont.Height = -16
    TitleFont.Name = #26032#32048#26126#39636
    TitleFont.Style = []
    OnDblClick = DBGrid1DblClick
    Columns = <
      item
        Expanded = False
        FieldName = 'DOC09'
        Visible = True
      end
      item
        Expanded = False
        FieldName = 'typedesc'
        Visible = True
      end
      item
        Expanded = False
        FieldName = 'childdesc'
        Visible = True
      end
      item
        Expanded = False
        FieldName = 'DOC02'
        Visible = True
      end
      item
        Expanded = False
        FieldName = 'DOC03'
        Visible = True
      end
      item
        Expanded = False
        FieldName = 'DOC04'
        Visible = True
      end
      item
        Expanded = False
        FieldName = 'DOC07'
        Visible = True
      end
      item
        Expanded = False
        FieldName = 'DOC08'
        Title.Caption = #20027#26088
        Visible = True
      end
      item
        Expanded = False
        FieldName = 'CREATEDATE'
        Visible = True
      end
      item
        Expanded = False
        FieldName = 'CREATEUSER'
        Visible = True
      end
      item
        Expanded = False
        FieldName = 'UPDATEDATE'
        Visible = True
      end
      item
        Expanded = False
        FieldName = 'UPDATEUSER'
        Visible = True
      end>
  end
  object DataSource1: TDataSource
    DataSet = queMain
    Left = 240
    Top = 288
  end
  object queMain: TADOQuery
    Connection = dm.gshank
    OnCalcFields = queMainCalcFields
    Parameters = <>
    SQL.Strings = (
      'select * from doc_file')
    Left = 352
    Top = 272
    object queMainDOC01: TWideStringField
      DisplayLabel = #36039#26009#32232#34399
      DisplayWidth = 16
      FieldName = 'DOC01'
    end
    object queMainDOC09: TWideStringField
      DisplayLabel = #36039#26009#26178#38291#35498#26126
      DisplayWidth = 20
      FieldName = 'DOC09'
      Size = 50
    end
    object queMaintypedesc: TStringField
      DisplayLabel = #22823#39006#35498#26126
      DisplayWidth = 20
      FieldKind = fkCalculated
      FieldName = 'typedesc'
      Calculated = True
    end
    object queMainchilddesc: TStringField
      DisplayLabel = #23567#39006#35498#26126
      DisplayWidth = 30
      FieldKind = fkCalculated
      FieldName = 'childdesc'
      Size = 40
      Calculated = True
    end
    object queMainDOC02: TWideStringField
      DisplayLabel = #36039#26009#24180#24230
      DisplayWidth = 8
      FieldName = 'DOC02'
      Size = 10
    end
    object queMainDOC03: TWideStringField
      DisplayLabel = #36039#26009#26376#20221
      FieldName = 'DOC03'
      Size = 2
    end
    object queMainDOC04: TWideStringField
      DisplayLabel = #36039#26009#23395#21029
      FieldName = 'DOC04'
      Size = 6
    end
    object queMainDOC05: TWideStringField
      DisplayLabel = #36039#26009#22823#39006
      FieldName = 'DOC05'
      Visible = False
      Size = 2
    end
    object queMainDOC06: TWideStringField
      DisplayLabel = #36039#26009#23567#39006
      FieldName = 'DOC06'
      Visible = False
      Size = 2
    end
    object queMainDOC07: TWideStringField
      DisplayLabel = #38651#23376#27284#26696
      DisplayWidth = 20
      FieldName = 'DOC07'
      Size = 100
    end
    object queMainDOC08: TWideStringField
      DisplayLabel = #20633#35387
      DisplayWidth = 20
      FieldName = 'DOC08'
      Size = 100
    end
    object queMainCREATEDATE: TDateTimeField
      DisplayLabel = #24314#31435#26085#26399
      DisplayWidth = 10
      FieldName = 'CREATEDATE'
    end
    object queMainCREATEUSER: TWideStringField
      DisplayLabel = #24314#31435#32773
      FieldName = 'CREATEUSER'
      Size = 10
    end
    object queMainUPDATEDATE: TDateTimeField
      DisplayLabel = #26356#26032#26085#26399
      DisplayWidth = 10
      FieldName = 'UPDATEDATE'
    end
    object queMainUPDATEUSER: TWideStringField
      DisplayLabel = #26356#26032#32773
      FieldName = 'UPDATEUSER'
      Size = 10
    end
  end
  object ADOQuery1: TADOQuery
    Connection = dm.gshank
    Parameters = <>
    Left = 336
    Top = 136
  end
  object IdFTP1: TIdFTP
    Host = 'web'
    Password = '04403568'
    User = 'webftp'
    Left = 496
    Top = 240
  end
end
