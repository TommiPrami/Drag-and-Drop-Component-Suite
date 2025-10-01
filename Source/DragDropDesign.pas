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

type
  TDataFormatNameEditor = class(TStringProperty)
  public
    function GetAttributes: TPropertyAttributes; override;
    procedure GetValues(Proc: TGetStrProc); override;
  end;

////////////////////////////////////////////////////////////////////////////////
//
//              Component and Design-time editor registration
//
////////////////////////////////////////////////////////////////////////////////
procedure Register;
begin
  RegisterPropertyEditor(TypeInfo(string), TDataFormatAdapter, 'DataFormatName',
    TDataFormatNameEditor);
  RegisterComponents(DragDropComponentPalettePage,
    [TDropEmptySource, TDropEmptyTarget, TDropDummy, TDataFormatAdapter,
    TDropFileTarget, TDropFileSource, TDropBMPTarget, TDropBMPSource,
    TDropMetaFileTarget, TDropImageTarget, TDropURLTarget, TDropURLSource,
    TDropPIDLTarget, TDropPIDLSource, TDropTextTarget, TDropTextSource,
    TDropComboTarget]);
  RegisterComponents(DragDropComponentPalettePage,
    [TDropHandler, TDragDropHandler, TDropContextMenu]);
end;

{ TDataFormatNameEditor }

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

var
  SplashScreen: HBITMAP;

initialization
  SplashScreen := LoadBitmap(hInstance, 'DRAGDROPSUITE');
  (SplashScreenServices as IOTasplashScreenServices).AddPluginBitmap('Drag and Drop Component Suite', SplashScreen, False, 'Open Source', 'Private build');
end.
