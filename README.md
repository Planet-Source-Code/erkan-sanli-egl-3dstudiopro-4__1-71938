<div align="center">

## EGL\_3DStudioPro 4

<img src="PIC2010811610551167.jpg">
</div>

### Description

An excellent 3DS viewer application. This 3D application allows you to load 3DS files into the application and view them in 3D and save bitmap. Also editing texture parameters; change, tiling, moving, mirroring, flipping, rotating, change opacity-transparency value. Realtime changing material color using RGB and HSL on colordialog. Total 14 different viewstyle. New styles; texturized wireframe, transparent (hidden) wireframe, photo realistic.

If you want to without GradientFill API use "modMyTriangleGradient.DrawTriangleGradientA". Also this sub filling transparent gradient triangle. Another sub

"modMyLineTexture.DrawLineTex". This sub drawing multicolored line. Pure vb, without using OpenGL or DirectX. Include gouraud shading, Delaunay triangulation, clipping, and other stuffs. (Zip:708 kb)
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |2010-08-11 12:06:06
**By**             |[Erkan Sanli](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/erkan-sanli.md)
**Level**          |Advanced
**User Rating**    |5.0 (95 globes from 19 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Graphics](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/graphics__1-46.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[EGL\_3DStud2186048112010\.zip](https://github.com/Planet-Source-Code/erkan-sanli-egl-3dstudiopro-4__1-71938/archive/master.zip)

### API Declarations

```
Private Declare Function CreateCompatibleDC Lib "gdi32" ...
Private Declare Function CreateDIBSection Lib "gdi32" ...
Private Declare Function SelectObject Lib "gdi32" ...
Private Declare Function DeleteObject Lib "gdi32" ...
Private Declare Function DeleteDC Lib "gdi32" ...
Private Declare Function BitBlt Lib "gdi32" ...
Private Declare Function StretchBlt Lib "gdi32" ...
Private Declare Function SetStretchBltMode Lib "gdi32" ...
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ...
Private Declare Sub ZeroMemory Lib "kernel32" Alias "RtlZeroMemory" ...
Private Declare Function VarPtrArray Lib "MSVBVM60.dll" Alias "VarPtr" ...
```





