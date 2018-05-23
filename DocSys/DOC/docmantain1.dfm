object fmbigtype: Tfmbigtype
  Left = 269
  Top = 135
  Width = 608
  Height = 480
  BorderIcons = [biSystemMenu]
  Caption = #20195#30908#22823#39006#32173#35703
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  Position = poScreenCenter
  PixelsPerInch = 96
  TextHeight = 13
  object DBNavigator1: TDBNavigator
    Left = 0
    Top = 0
    Width = 600
    Height = 25
    DataSource = DataSource1
    Align = alTop
    TabOrder = 0
  end
  object DBGrid1: TDBGrid
    Left = 0
    Top = 25
    Width = 600
    Height = 428
    Align = alClient
    DataSource = DataSource1
    Font.Charset = ANSI_CHARSET
    Font.Color = clWindowText
    Font.Height = -16
    Font.Name = #26032#32048#26126#39636
    Font.Style = []
    ParentFont = False
    TabOrder = 1
    TitleFont.Charset = ANSI_CHARSET
    TitleFont.Color = clWindowText
    TitleFont.Height = -16
    TitleFont.Name = #26032#32048#26126#39636
    TitleFont.Style = []
    Columns = <
      item
        Expanded = False
        FieldName = 'BIG01'
        Visible = True
      end
      item
        Expanded = False
        FieldName = 'BIG02'
        Visible = True
      end>
  end
  object DataSource1: TDataSource
    DataSet = ADOQuery1
    Left = 128
    Top = 128
  end
  object ADOQuery1: TADOQuery
    Connection = dm.gshank
    Parameters = <>
    SQL.Strings = (
      'select * from doc_big_type')
    Left = 280
    Top = 136
    object ADOQuery1BIG01: TWideStringField
      DisplayLabel = #22823#39006#20195#30908
      FieldName = 'BIG01'
      Size = 2
    end
    object ADOQuery1BIG02: TWideStringField
      DisplayLabel = #22823#39006#35498#26126
      FieldName = 'BIG02'
      Size = 50
    end
  end
end
