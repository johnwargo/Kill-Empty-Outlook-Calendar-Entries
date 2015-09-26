{ ****************************************************************************
  Kill Calendar

  The software that syncs my Outlook calendar to my Android device was
  complaining that it found empty calendar entries, so I wrote this program
  to delete all of them

  Learned how to work with the Outlook calendar from Delphi using:
  http://www.scalabium.com/faq/dct0128.htm

  Learned how to access Outlook items using:
  http://www.scalabium.com/faq/dct0121.htm

  John M. Wargo
  September 25, 2015
  **************************************************************************** }
unit kc_main;

interface

uses
  ComObj, Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants,
  System.Classes,

  Vcl.Graphics, Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.ComCtrls,
  Vcl.StdCtrls;

type
  TfrmMain = class(TForm)
    StatusBar1: TStatusBar;
    output: TMemo;
    procedure FormActivate(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmMain: TfrmMain;

implementation

{$R *.dfm}

procedure TfrmMain.FormActivate(Sender: TObject);
const
  olAppointmentItem = $00000001;

var
  outlook, oiItem, ns, folder: OLEVariant;
  i, intFolderType, numItems: Integer;
  msgText, subject: String;

  { to find a default Calendar folder }
  function GetCalendarFolder(folder: OLEVariant): OLEVariant;
  var
    i: Integer;
  begin
    for i := 1 to folder.Count do
    begin
      if (folder.Item[i].DefaultItemType = olAppointmentItem) then
        result := folder.Item[i]
      else
        result := GetCalendarFolder(folder.Item[i].Folders);
      if not VarIsNull(result) and not VarIsEmpty(result) then
        break;
    end;
  end;

begin
  output.Lines.add('Creating Outlook Object');
  // initialize a connection to Outlook
  outlook := CreateOLEObject('Outlook.Application');
  // get the MAPI namespace
  ns := outlook.GetNamespace('MAPI');
  // get a default Contacts folder
  folder := GetCalendarFolder(ns.Folders);
  output.Lines.add('Default calendar folder: ' + string(folder));
  // if  Calendar folder is found
  if VarIsNull(folder) and not VarIsEmpty(folder) then
  begin
    // Then tell the user
    msgText := 'Unable to determine default calendar folder';
    ShowMessage(msgText);
    output.Lines.add(msgText);
  end else begin
    // Process entries in the folder
    output.Lines.add(Format('Searching "%s" folder', [folder]));
    intFolderType := folder.DefaultItemType;
    numItems := folder.Items.Count;
    if (numItems > 0) then
    begin
      output.Lines.add(Format('Found %d items', [numItems]));
      // Have to go backwards through the list to delete them apparently
      for i := numItems downto 1 do
      begin
        oiItem := folder.Items[i];
        subject := oiItem.subject;
        if (Length(subject) < 1) then
        begin
          msgText := Format('Deleting item %d', [i]);
          output.Lines.add(msgText);
          // Uncomment the following line when you're ready to delete entries
          // oiItem.Delete;
        end;
      end;
      output.Lines.add('All done!');
    end else begin
      output.Lines.add('No calendar entries found');
    end;
  end;
end;

end.
