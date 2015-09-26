object frmMain: TfrmMain
  Left = 0
  Top = 0
  Caption = 'Delete Empty Calendar Entries'
  ClientHeight = 461
  ClientWidth = 738
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  OnActivate = FormActivate
  PixelsPerInch = 96
  TextHeight = 13
  object StatusBar1: TStatusBar
    Left = 0
    Top = 442
    Width = 738
    Height = 19
    Panels = <>
    SimplePanel = True
    SimpleText = 'Copyright 2015 McNelly SoftWorks, LLC'
  end
  object output: TMemo
    Left = 0
    Top = 0
    Width = 738
    Height = 442
    Align = alClient
    ReadOnly = True
    ScrollBars = ssVertical
    TabOrder = 1
  end
end
