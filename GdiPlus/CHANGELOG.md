## 0.9.38
- Add easier to use variations of DrawDriverString and MeasureDriverString.
- Factor out VBGraphics to its own class.
- Add GpRect[F] constructors that default size to 0. Clean up coordinate types.
- Accept null strings for drawing and measuring. They are treated like empty strings.
- Add ArrayPtr and refactor all locations where a pointer to array data is needed.
- Fix compilation errors on win64.

## 0.9.37
- Fix bugs:
  - Fix GenericDefaultStringFormat and GenericTypographicStringFormat returning uninitialized objects.
  - Fix StringFormat.SetTabStops to work correctly with arrays that start at an index other than 1.
  - Fix StringFormat constructor not setting the native handle.
- Add StopOnErrors strategy.
- Ensure that Graphics.MeasureCharacterRanges correctly outputs the result array, even if none was provided.
- Add a MeasureCharacterRanges overload that returns the region array.

## 0.9.36

- Fix bugs:
  - Fix StringFormat.TabStops property.
  - Fix GpStringFormat.SetTabStops2 signature.
  - Fix non-functional InstalledFontCollection, PrivateFontCollection, Matrix, PathIterator and Region classes
  - Fix memory leak in PrivateFontCollection by fixing the signature of GpPrivateFontCollection.Delete.
  - Fix Graphics.MeasureStringS failing to build due to incorrect use of IsMissing on a UDT.
  - Fix Matrix.Reset incorrectly zeroing out the matrix. Instead it sets the matrix to identity.

## 0.9.35

- Fix bugs:
  - Initialization of the VB compatibility layer in Graphics.
  - Building without WinDevLib.
  - Missing ByVal for GpPoint\[F], causing crashes.
  - Missing (GpRect, rgb, rgb) constructor in LinearGradientBrush.
- Add the global error handling strategy, and a global error handler delegate.
- Further uniformize Graphics, Matrix, LinearGradientBrush, PathGradientBrush, TextureBrush, Color, and FontFamily.
- Make PathData inherit GdiPlusBase.
- Use byte-sized PathPointType instead of the generic Byte type.
- Add a preliminary manual to the repository.
- Add the missing changelog entry for 0.9.34.

## 0.9.34

- Fix bugs:
  - PathGradientBrush: SurroundColorCount returning a wrong count, RotateTransform having a messed up name
  - StringFormat: type returned by the Trimming property
  - GraphicsPath: native function used in IsOutlineVisible with integer arguments
  - Pen: incorrect types in DashOffset property
  - Image: clone not working
  - LineCap: BaseInset returning a wrong value
  - GdipEffects: ColorLutParams constructor incorrectly referencing the arguments

- Make the GpXxx (UDT-based) and Xxx (Class-based) APIs consistent across all classes/types.

  All of the functionality built on top of the native API is implemented in UDT-based types.
  The class types use the UDT API. This makes it as easy to use the UDT-based API as the Class-based
  API. Note that the UDT API does not automatically manage resources due to limitations of the UDT
  lifecycle methods in twinBASIC.
  
- Make the Imaging-related types and classes easier to use.

- Remove the reference+count APIs from the classes. The supported way of passing around arrays is
  to use native tB arrays.
  
- Define ITransformable and implement it on all classes that can be transformed.

  All transformable objects now support RotateAt and Shear even if the native API doesn't.
  
- Add BufferedGraphicsXxx constructors to the Graphics type to make it easy to perform buffered painting
  of windows and controls without flicker.
  
- Rename free-standing Graphics constructor functions to prevent silent misuse bugs.

- Remove the Brush class. The Brush type is an alias of the IBrush interface.

- Make LastStatus private in GdiPlusBase. This forces SetStatus to be used, and makes it a single point
  where all errors can be caught.

## 0.9.33

- Fix Font.InitMetrics returning a wrong value.
- Fix GraphicsPath.New incorrectly invoking PathData property rather than GdipPathData.PathData constructor.