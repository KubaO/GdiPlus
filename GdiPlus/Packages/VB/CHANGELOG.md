# Change Log
All notable changes to "twinBASIC WinNativeForms" will be documented in this file.

## [v0.0.31.0, 15th September 2022]
- added: PictureBox
- improved: made changes to ensure nothing within the package is being exposed unnecessarily

## [v0.0.30.0, 11th September 2022]
- added: FileListBox

## [v0.0.29.0, 10th September 2022]
- fixed: implemented HandleEraseBackground for ComboBox [ https://github.com/twinbasic/twinbasic/issues/1148 ]

## [v0.0.28.0, 9th September 2022]
- added: DriveListBox control
- added: DirListBox control

## [v0.0.27.0, 8th September 2022]
- added: TreeView control

## [v0.0.26.0, 31st August 2022]
- improved: added error handling to some Ambient ActiveX properties for better ActiveX control support

## [v0.0.25.0, 25th May 2022]
- fixed: Form.Load event now fires after the root constructor [ https://github.com/WaynePhillipsEA/twinbasic/issues/799 ]

## [v0.0.24.0, 22nd April 2022]
- fixed: typo CombBox.twin -> ComboBox.twin

## [v0.0.23.0, 22nd April 2022]
- added: UserControl

## [v0.0.22.0, 7th April 2022]
- added: event Form.Unload

## [v0.0.21.0, 5th April 2022]
- fixed: BaseForm.Refresh was missing flag RDW_ERASE in call to RedrawWindow

## [v0.0.20.0, 2nd April 2022]
- added: property HScrollBar._Default [DefaultMember]
- added: property VScrollBar._Default [DefaultMember]
- added: property Timer._Default [DefaultMember]
- added: property TextBox._Default [DefaultMember]
- added: property OptionButton._Default [DefaultMember]
- added: property ListBox._Default [DefaultMember]
- added: property Label._Default [DefaultMember]
- added: property ComboBox._Default [DefaultMember]
- added: property CheckBox._Default [DefaultMember]

## [v0.0.19.0, 1st April 2022]
- changed [DefaultDesignerEvent] of Form from 'Activate' event to 'Load' event
- BaseForm.WindowState now works at runtime

## [v0.0.18.0, 24th February 2022]
- refactored: WindowsAPI to avoid naming conflicts

## [v0.0.17.0, 19th February 2022]
- minor tweak for Image EraseBackground routine

## [v0.0.16.0, 19th February 2022]
- changed: all toolbox images, kindly provided by @Robin1997#8689 over on Discord.  Thanks Rob!
- fixed: Form.hWnd property is now exposed [ https://github.com/WaynePhillipsEA/twinbasic/issues/706#issuecomment-1044075841 ]
- added: Image.Border property [ https://github.com/WaynePhillipsEA/twinbasic/issues/706#issuecomment-1045270615 ]
- fixed: changing the Image.Picture property at runtime now triggers a Refresh()

## [v0.0.15.0, 18th February 2022]
- added: initial implementation of Image control

## [v0.0.14.0, 17th February 2022]
- fixed: WM_SETFONT handling to match VB6 [ https://github.com/WaynePhillipsEA/twinbasic/issues/706#issuecomment-1042747336 ]
- improved: changed CB_GETCOMBOBOXINFO to use WindowsAPI.GetComboBoxInfo for Win2K support [ https://github.com/WaynePhillipsEA/twinbasic/issues/706#issuecomment-1042674289 ]
- fixed: changed ListBox.DblClick event to fire from LBN_DBLCLK rather than WM_LBUTTONDBLCLK [ https://github.com/WaynePhillipsEA/twinbasic/issues/706#issuecomment-1042670789 ] 
- reverted: IBeam change from v0.0.12.0, and fixed the WM_SETCURSOR handling [ https://github.com/WaynePhillipsEA/twinbasic/issues/706#issuecomment-1042665044 ]
- improved: added ListBox item checkbox padding to better match VB6 [ https://github.com/WaynePhillipsEA/twinbasic/issues/706#issuecomment-1042801078 ] 
- fixed: ProgressBar compile error on x64

## [v0.0.13.0, 17th February 2022]
- added: ProgressBar control
- added: SubClass optional argument to CreateRootWindowElement base functions

## [v0.0.12.0, 16th February 2022]
- fixed: TextBox and ComboBox MousePointer now defaults to IBeam

## [v0.0.11.0, 16th February 2022]
- renamed: package to WinNativeForms
- fixed: ListBox events not firing due to missing LBS_NOTIFY
- added: ListBox support for vbListBoxCheckbox (checked listbox)
- added: Scroll event support to ComboBox, ListBox, HScrollBar and VScrollBar
- added: internal event of MouseWheel(Delta)
- added: internal event of PreMouseDown(Buttons, Shift, X, Y)
- added: internal event of PreProcessMessage(Msg, wParam, lParam, MutedReturnValue)
- added: internal event of PostProcessMessage(Msg, wParam, lParam)
- added: ListBox.MaxCheckboxSize property (Long)
- added: ListBox.WheelScrollEvent property (Boolean)
- added: ComboBox.WheelScrollEvent property (Boolean)

## [v0.0.10.0, 14th February 2022]
- TextBox MultiLine mode fixes [ https://github.com/WaynePhillipsEA/twinbasic/issues/734 ]
- other minor source changes

## [v0.0.9.0, 9th February 2022]
- changed WindowsAPI.SetFont, as directed by @Kr00l [ https://github.com/WaynePhillipsEA/twinbasic/issues/706#issuecomment-1033828143 ]
- runtime font changes now trigger a Refresh()

## [v0.0.8.0, 8th February 2022]
- removed Timer constructor attempting to override the Width/Height properties, causing an unhandled error

## [v0.0.7.0, 6th February 2022]
- tidy up and better use of Implements-Via inheritance for BaseControl

## [v0.0.5.0, 5th February 2022]
- added an Initialize() event to all controls, to help with control inheritance

## [v0.0.4.0, 5th February 2022]
- added simple New() constructors for each control, to allow for control inheritance

## [v0.0.3.0, 5th February 2022]
- fixed this change log
- changed 'DisableVisualStyles' property to 'VisualStyles'

## [v0.0.2.0, 4th February 2022]
- initial release
