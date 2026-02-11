# GDI+ for twinBASIC

This is a preview of a higher-level GDI+ interface for twinBASIC. It is modeled after the C++ wrapper that is the official GDI+ API from Microsoft.

To get the `.twinpack` package file, open the project and build it.

The project is also a public package in TwinBasic's online package library.

There is a [basic manual](Documentation/index.md) available.

TO DO:

- [ ] \+ more glue/convenience code



![Screenshot of the Demo showing "Welcome to GDI+" text on a background consisting of: a diagonal light green line over a diagonal cyan-to-blue gradient (top left to bottom right).](Documentation/Images/trivial-demo.png)



The window above is produced by the following code:

``` vb
Class Form1
    Dim mGdiPlus As GdiPlusUser
    
    Sub Form_Load()
        ScaleMode = vbPixels
        Set mGdiPlus = GdiPlusUser()
    End Sub
            
    Sub Paint() Handles Form.Paint, Form.Resize
        Dim rect As Any = GpRect(0, 0, ScaleWidth, ScaleHeight)
        Dim gr As Any = BufferedGraphicsFromHDC(hDC, rect)
        Dim br As Any = LinearGradientBrush(rect, Cyan, Blue)
        Dim ft As Any = GdiPlus.Font("Segoe UI", 20)
        Dim fbr As Any = SolidBrush(Black)
                        
        gr.FillRectangle(br, rect)
        gr.DrawLine(Pen(LightGreen, 3.0), rect)
        gr.DrawString("Welcome to GDI+", ft, GpPointF(50, 50), fbr)
    End Sub
End Class
```

Note:

- An instance of `GdiPlusUser` must exist while GDI+ objects are used. An arbitrary number of instances can exist at one time, as long as *any* exist, the GDI+ library is kept active. Since GDI+ is typically used to render forms, or is triggered by actions in the form-based UI, it is sufficient to add an instance of GdiPlusUser to any Form or other class whose code used GDI+.
- All classes can be instantiated either with `New AClass(...)` or using a convenience "constructor" function with the same signature, i.e. `AClass(...)`. This reduces verbosity in the user code.
- UDTs such as `GpPointF`, `GpRectF` etc. also have constructor functions.

