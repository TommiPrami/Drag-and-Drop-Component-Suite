unit DragDropDesign;

(*
 * Drag and Drop Component Suite
 *
 * Copyright (c) Angus Johnson & Anders Melander
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at https://mozilla.org/MPL/2.0/.
 *
 *)

// -----------------------------------------------------------------------------
//
// Contains design-time support for the drag and drop components.
//
// -----------------------------------------------------------------------------
// TODO : Default event for target components should be OnDrop.
// TODO : Add parent form to Target property editor list.

interface

{$include DragDrop.inc}

procedure Register;

implementation

uses
  System.Classes,
  System.SysUtils,
  Windows,
  DesignIntf,
  DesignEditors,
  ToolsAPI,
  DragDrop,
  DropSource,
  DropTarget,
  DragDropFile,
  DragDropGraphics,
  DragDropContext,
  DragDropHandler,
  DropHandler,
  DragDropInternet,
  DragDropPIDL,
  DragDropText,
  DropComboTarget;

////////////////////////////////////////////////////////////////////////////////
//
//      Property editor for the TDataFormatAdapter.DataFormatName property
//
////////////////////////////////////////////////////////////////////////////////
type
  TDataFormatNameEditor = class(TStringProperty)
  public
    function GetAttributes: TPropertyAttributes; override;
    procedure GetValues(Proc: TGetStrProc); override;
  end;

function TDataFormatNameEditor.GetAttributes: TPropertyAttributes;
begin
  Result := [paValueList, paSortList, paMultiSelect];
end;

procedure TDataFormatNameEditor.GetValues(Proc: TGetStrProc);
var
  i: Integer;
begin
  for i := 0 to TDataFormatClasses.Count-1 do
    Proc(TDataFormatClasses.Formats[i].ClassName);
end;


////////////////////////////////////////////////////////////////////////////////
//
//      Component and Design-time editor registration
//
////////////////////////////////////////////////////////////////////////////////
procedure Register;
begin
  RegisterPropertyEditor(TypeInfo(string), TDataFormatAdapter, 'DataFormatName',
    TDataFormatNameEditor);

  RegisterComponents(DragDrop.DragDropComponentPalettePage,
    [TDropEmptySource, TDropEmptyTarget, TDropDummy, TDataFormatAdapter,
    TDropFileTarget, TDropFileSource, TDropBMPTarget, TDropBMPSource,
    TDropMetaFileTarget, TDropImageTarget, TDropURLTarget, TDropURLSource,
    TDropPIDLTarget, TDropPIDLSource, TDropTextTarget, TDropTextSource,
    TDropComboTarget]);

  RegisterComponents(DragDrop.DragDropComponentPalettePage,
    [TDropHandler, TDragDropHandler, TDropContextMenu]);
end;


////////////////////////////////////////////////////////////////////////////////
//
//      Branding
//
////////////////////////////////////////////////////////////////////////////////
const
  sProductName = 'Drag and Drop Component Suite';
  sProductVersion = '6.x';
  sProductCopyright = 'Copyright '#$00A9' 1997-%s'#13#10'Angus Johnson & Anders Melander'#13#10'All rights reserved.';
  sProductLogo = 'DRAGDROPSUITE';
  sProductLicense = 'Open Source: Mozilla Public License, v. 2.0';
  sProductSKU = 'private release'; // Will be appended, with a space before, to the product name on the splash
  sProductRegistered = True;


////////////////////////////////////////////////////////////////////////////////
//
//      IDE splash screen
//
////////////////////////////////////////////////////////////////////////////////
type
  SplashInfo = record
  private
    class var FLogoBitmap: HBITMAP;
  public
    class procedure Initialize; static;
    class procedure Finalize; static;
  end;

class procedure SplashInfo.Initialize;
var
  Title: string;
begin
  FLogoBitmap := LoadBitmap(hInstance, sProductLogo);
  if (sProductVersion <> '') then
    Title := Format('%s %s', [sProductName, sProductVersion])
  else
    Title := sProductName;
  (SplashScreenServices as IOTasplashScreenServices).AddPluginBitmap(Title, FLogoBitmap, not sProductRegistered, sProductLicense, sProductSKU);
end;

class procedure SplashInfo.Finalize;
begin
  DeleteObject(FLogoBitmap);
end;


////////////////////////////////////////////////////////////////////////////////
//
//      IDE about box
//
////////////////////////////////////////////////////////////////////////////////
type
  AboutInfo = record
  private
    class var FLogoBitmap: HBITMAP;
    class var FAboutInfoIndex: integer;
  public
    class procedure Initialize; static;
    class procedure Finalize; static;
  end;

class procedure AboutInfo.Initialize;
var
  AboutBoxServices: IOTAAboutBoxServices;
  Title, Info, Copyright: string;
begin
  FAboutInfoIndex := -1;

  if Supports(BorlandIDEServices, IOTAAboutBoxServices, AboutBoxServices) then
  begin
    FLogoBitmap := LoadBitmap(hInstance, sProductLogo);

    if (sProductVersion <> '') then
      Title := Format('%s %s', [sProductName, sProductVersion])
    else
      Title := sProductName;
    Copyright := Format(sProductCopyright, [FormatDateTime('YYYY', Now)]);
    Info := Format('%s %s'#13#10#13#10'%s', [sProductName, sProductVersion, Copyright]);

    FAboutInfoIndex := AboutBoxServices.AddPluginInfo(Title, Info, FLogoBitmap, not sProductRegistered, sProductLicense, sProductSKU);
  end;
end;

class procedure AboutInfo.Finalize;
var
  AboutBoxServices: IOTAAboutBoxServices;
begin
  if (FAboutInfoIndex <> -1) and Supports(BorlandIDEServices, IOTAAboutBoxServices, AboutBoxServices) and (FAboutInfoIndex <> -1) then
  begin
    AboutBoxServices.RemovePluginInfo(FAboutInfoIndex);

    DeleteObject(FLogoBitmap);
  end;
end;

////////////////////////////////////////////////////////////////////////////////

initialization
  SplashInfo.Initialize;
  AboutInfo.Initialize;

finalization
  SplashInfo.Finalize;
  AboutInfo.Finalize;

end.
