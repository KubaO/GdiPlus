# GdiPlus for TwinBasic - Development Guide

## Scope

**Only source files under `GdiPlus/Sources/` and documentation under `Documentation` is in scope for this project.** Do not modify or suggest changes to source files in `Samples/`, `Experiments/`, or any other directory outside `GdiPlus/Sources/`.

---

## Project Overview

**GdiPlus** is a comprehensive TwinBasic package that provides bindings to the Windows GDI+ graphics library. GDI+ is a successor to the GDI (graphics device interface) API and provides vector graphics, imaging, and typography capabilities.

The package is licensed under the MIT License and is maintained at https://github.com/kubao/GdiPlus

---

## Repository Structure

### Main Project
- **`GdiPlus.twinproj`** - Primary TwinBasic project file; builds both the library and a small demo EXE
- **`GdiPlus/Sources/`** - Main source code directory
  - **`GdiPlus/`** - Core GdiPlus module implementations (~35 files)
  - **`GdiPlusTest.twin`** - Demo EXE entry point (analog clock window using raw WndProc)
  - **`Tests/`** - Compile-time tests for type API overload resolution

### Documentation
- **`Documentation/`** - User documentation
  - **`index.md`** - Main documentation with class reference
  - **`Images/`** - Documentation assets

---

## Core TwinBasic Modules

### Platform & Base Types
- **`Platform.twin`** - Raw GDI+ C-style structs: `GDIP_UUID`, `GDIP_POINT`, `GDIP_RECT`, `GDIP_RECTF`, `GDIP_SIZE`, `GDIP_BOOL`
- **`Types.twin`** - TwinBasic-specific types and wrappers: `GpPoint[F]`, `GpSize[F]`, `GpRect[F]` UDTs with methods; `GpStatus` enum; callback delegates; `ITransformable`, `IDeviceContext`, `GdiPlusAbort` interfaces
- **`Native.twin`** - Internal module providing `GetNative()` overloaded accessors for extracting native handles from GdiPlus wrapper objects (used internally, not part of the public API)

### Graphics & Rendering
- **`GdiPlusBase.twin`** - Base class for all GdiPlus objects; provides `LastResult`/`GetLastResult`, error handling hooks, and GdipAlloc/Free
- **`GdiPlusUser.twin`** - GDI+ initialization and lifetime management; also defines `GdiplusStartupInput`, `GdiplusStartupOutput`, startup param enums
- **`Graphics.twin`** (133KB) - Main rendering class; creates Graphics objects from HDC, HWND, or Image; provides all drawing methods
- **`VBGraphics.twin`** - Visual Basic compatible graphics interface
- **`Implementation.twin`** - Private internal module with `ArrayLen`, `ArrayPtr`, `SetStaticSA`, `StaticSA` helpers

### Geometric Primitives (UDTs with methods in `Types.twin`)
- **`GpPointF`** / **`GpPoint`** - Floating-point / integer 2D coordinates; support Add, Sub, Scale, Normalize, Transpose, DotProd, Sum, Diff, Prod, Quot
- **`GpSizeF`** / **`GpSize`** - Floating-point / integer dimensions; support Add, Sub, Scale, Sum, Diff, Prod, Quot
- **`GpRectF`** / **`GpRect`** - Floating-point / integer rectangles; support Offset, Inflate, Intersect, Unite, Contains, conversion methods
- **`Matrix.twin`** - Affine transformation matrix operations
- **`GraphicsPath.twin`** (43KB) - Path construction and manipulation
- **`PathData.twin`** - Path data structures
- **`PathIterator.twin`** - Path iteration and enumeration
- **`Region.twin`** - Region clipping and hit testing

### Brushes & Pens
- **`Brush.twin`** - `IBrush` interface (with `Alias Brush As IBrush`); `GpBrush` native handle type; `BrushFrom` factory
- **`SolidBrush.twin`** - Solid color brushes
- **`HatchBrush.twin`** - Hatch pattern brushes
- **`LinearGradientBrush.twin`** (26KB) - Linear gradient brushes
- **`PathGradientBrush.twin`** (20KB) - Path-based gradient brushes
- **`TextureBrush.twin`** - Bitmap texture brushes
- **`Pen.twin`** (20KB) - Pen styling and line drawing

### Colors
- **`Color.twin`** (19KB) - `Color` UDT (ARGB, with A/R/G/B properties, HSL conversion, IsNull); `CColor` enum with named constants (e.g., `CColor.Red`, `CColor.White`); factory functions `Color(r, g, b)`, `Color(a, r, g, b)`, `Color(argb)`, `NoColor()`
- **`ColorMatrix.twin`** - `ColorMatrix` type (5x5 Single matrix), `ColorChannelLUT`, `HistogramFormat` enum, `ColorMatrixFlags`, `ColorAdjustType`

### Typography
- **`Font.twin`** - Font object and font properties
- **`FontFamily.twin`** - Font family definitions
- **`FontCollection.twin`** - Font collection management
- **`InstalledFontCollection.twin`** - System installed fonts
- **`PrivateFontCollection.twin`** - Custom font collections
- **`StringFormat.twin`** (33KB) - Text formatting and layout options

### Image & Bitmap Operations
- **`Image.twin`** (30KB) - Base image class and operations
- **`Bitmap.twin`** (30KB) - Bitmap-specific operations and properties
- **`Metafile.twin`** (30KB) - Vector graphic metafile support
- **`CachedBitmap.twin`** - Cached bitmap optimization
- **`ImageAttributes.twin`** - Image color and effect attributes
- **`ImageCodec.twin`** - Image codec information
- **`Imaging.twin`** (46KB) - Image encoding/decoding and metadata
- **`PixelFormats.twin`** - Pixel format definitions

### Effects & Filters
- **`Effects.twin`** (26KB) - Image effects and filters
  - Blur, Bright/Contrast, ColorBalance
  - ColorCurve, GaussianBlur, HueSaturationLightness
  - Levels, RedEyeRemoval, Sharpen, Tint

### Line & Cap Styles
- **`LineCap.twin`** - Line cap styling
- **`ArrowCap.twin`** - Arrow cap definitions
- **`MetaHeader.twin`** - Metafile header structures

### Enumerations
- **`Enums.twin`** (56KB) - All GDI+ enumerations (FillMode, SmoothingMode, InterpolationMode, PixelFormat, FontStyle, HatchStyle, DashStyle, WrapMode, CompositingMode, etc.)

---

## Key Classes & Interfaces

### Initialization & Lifecycle
- **`GdiPlusUser`** - Must be instantiated to activate GDI+; manages library initialization and cleanup; reference-counted so multiple instances are safe; auto-selects GDI+ version 1 or 2 based on OS
- **`GdiPlusBase`** - Base class for all GdiPlus objects; provides `LastResult`, `GetLastResult`, `ClearResult`, `Status`, `StatusNL`

### Graphics Operations
- **`Graphics`** - Principal class for rendering; created from HDC, HWND, or Image
- **`GraphicsPath`** - Constructs and manipulates vector paths
- **`Region`** - Defines clipping and hit-test regions
- **`Matrix`** - Affine transformations (scale, rotate, translate, shear)

### Geometric Types (TwinBasic UDTs, not COM classes)
These are `Type` definitions with methods, not COM classes. They are value types (passed by value, no `Set`/`Nothing`).
- **`GpPoint`** / **`GpPointF`** - Integer and floating-point coordinates
- **`GpSize`** / **`GpSizeF`** - Integer and floating-point dimensions
- **`GpRect`** / **`GpRectF`** - Integer and floating-point rectangles

### Painting
- **`Pen`** - Line drawing with styles, caps, joins
- **`Brush`** / **`IBrush`** - Interface for all brush types (`Brush` is an alias for `IBrush`)
- **`SolidBrush`** - Solid colors
- **`LinearGradientBrush`** - Linear gradients
- **`PathGradientBrush`** - Radial/path-based gradients
- **`TextureBrush`** - Bitmap fills
- **`HatchBrush`** - Hatch patterns

### Images & Bitmaps
- **`Image`** - Base class for all image types
- **`Bitmap`** - Raster image operations
- **`Metafile`** - Vector graphic metafiles
- **`CachedBitmap`** - Pre-rendered bitmap caching

### Text & Fonts
- **`Font`** - Font definition with family, size, style
- **`FontFamily`** - Font family properties and metrics
- **`FontCollection`** - Base class for font collections
- **`InstalledFontCollection`** - System fonts
- **`PrivateFontCollection`** - Custom font loading
- **`StringFormat`** - Text layout, alignment, trimming, tabs
- **`VBGraphics`** - VB6-compatible text drawing interface

### Color & Effects
- **`Color`** - ARGB color UDT with A/R/G/B properties and HSL support; named constants via `CColor` enum
- **`ColorMatrix`** - 5x5 color transformation matrix UDT
- **`ImageAttributes`** - Color remapping, gamma, threshold, and transparency adjustments
- **`Effects`** - Image filters and effects

---

## Error Handling

All GdiPlus objects inherit from `GdiPlusBase` and store the last GDI+ status. The global `GdipErrorHandlingStrategy` (in module `GdipModule`) controls behavior on errors:

| Strategy | Behavior |
|---|---|
| `IgnoreErrors` | Errors stored in `LastResult` only (default) |
| `RaiseErrors` | Calls `Err.Raise` with code `GdipErrorCodeBase + GpStatus` |
| `HandleErrors` | Calls the `GdipErrorHandler` callback if set |
| `StopOnErrors` | Executes `Stop` for debugging |

Check `obj.LastResult` (non-clearing) or `obj.GetLastResult()` (clears to `Ok`) after operations.

---

## Testing

- **`GdiPlus/Sources/GdiPlusTest.twin`** - Demo EXE entry point module; creates an analog clock window using raw WndProc to exercise the library
- **`GdiPlus/Sources/Tests/Types.twin`** - Compile-time tests (`[DebugOnly]` subs) that verify overload resolution for `GpPoint[F]`, `GpSize[F]`, `GpRect[F]` API

---

## Version History

Current version: **0.9.0.43** (as of March 2026)

### Recent Changes
- **0.9.0.43** - Expand members and conversions for GpPoint[F], GpSize[F], GpRect[F] (Breaking: Add/Sub are now modifying; use Sum/Diff instead)
- **0.9.0.42** - Fix NoColor constructor, add Color.IsNull
- **0.9.0.41** - Fix invalid Gdip import names
- **0.9.0.40** - Fix ArrayLen, ArrayPtr functions
- **0.9.0.39** - Drop WinDevLib dependency (35x size reduction: 780KB)

See `GdiPlus/CHANGELOG.md` for complete version history.

---

## Using the Library

### Minimal Example
```vb
Class MyForm
    Dim mGdiPlus As GdiPlusUser

    Sub Form_Load()
        ScaleMode = vbPixels
        Set mGdiPlus = GdiPlusUser()   ' factory function (convention)
    End Sub

    Sub Form_Paint()
        Dim gr As Graphics
        Set gr = GraphicsFromHWND(Me.hWnd)

        Dim brush As SolidBrush
        Set brush = SolidBrush(Color(255, 0, 0))  ' Red

        gr.FillRectangle brush, 10, 10, 100, 100
        Set gr = Nothing
    End Sub
End Class
```

### Key Requirements
1. Create a `GdiPlusUser` instance using the factory function `GdiPlusUser()` (convention; `New GdiPlusUser()` also works)
2. `Graphics` objects are transient - create, use, and destroy them quickly
3. Prefer buffered graphics (`BufferedGraphicsFromHWND`, `BufferedGraphicsFromHDC`) for flicker-free drawing
4. Always set forms to `ScaleMode = vbPixels` before using GDI+
5. `GpPoint`, `GpRect`, etc. are value types (UDTs) — no `Set`/`Nothing`, passed by value

---

## Development Notes

- **Language**: TwinBasic (VB6-like syntax with modern features)
- **Bindings**: Wraps the native Windows GDI+ library (gdiplus.dll); P/Invoke declarations are embedded directly in each module
- **Dependencies**: VBA and VBRUN compatibility packages
- **Project Format**: TwinBasic project files (.twinproj) are binary XML-based project definitions
- **File Extensions**:
  - `.twinproj` - TwinBasic project file
  - `.twin` - TwinBasic source module

### Architecture Principles
- **Object-oriented design** - COM classes wrap native GDI+ handles
- **Resource management** - Objects implement `Class_Terminate` for cleanup
- **Native handle types** - Each class has a companion `Type` (e.g., `GpBrush`, `GpGraphics`, `GpToken`) that holds the native `LongPtr` handle and embeds `Declare` P/Invoke members using `ByVal Me` as the first argument, making native calls look like method calls on the handle
- **Factory functions** - Each class has a corresponding module-level factory function (e.g., `SolidBrush(...)`, `Pen(...)`, `Color(r, g, b)`) rather than exposing constructors directly
- **VB compatibility** - Provides familiar VB6-like interfaces alongside native GDI+ features
- **Modular structure** - Each GDI+ concept (brushes, fonts, effects, etc.) in separate files
- **1-based arrays** - `Types.twin` and `Matrix.twin` use `Option Base 1`; `StaticSA` arrays also use `lLbound = 1`

---

## Important Breaking Changes

### Version 0.9.0.43
The `Add` and `Sub` methods on geometric types (GpPoint[F], GpSize[F], GpRect[F]) changed from non-modifying to modifying. Use `Sum` and `Diff` instead for non-modifying operations.

---

## Package Dependencies

- **VBA** - VBA standard library compatibility
- **VBRUN** - Visual Basic runtime compatibility

Both are included as sub-packages within the main GdiPlus.twinproj.

---

## Additional Resources

- **GitHub Repository**: https://github.com/kubao/GdiPlus
- **Documentation**: `Documentation/index.md` - Comprehensive class and method reference
- **License**: MIT License - See repository for full license text
- **Author/Maintainer**: Sunderland Ober Consulting (c) 2026
