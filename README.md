<div align="center">

## Scale / Resize Graphics While Printing


</div>

### Description

To print images stored in any random sizes within the specified Height and Width.

Without selecting and applying the proper "ScaleFactor" the images tend to get distorted - squashed or stretched - when they are resized to be printed on the form, picture box or the printer object.

The code below will ensure that the image is not distorted nor croped at the time of printing.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Gajendra S\. Dhir](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/gajendra-s-dhir.md)
**Level**          |Beginner
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Graphics](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/graphics__1-46.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/gajendra-s-dhir-scale-resize-graphics-while-printing__1-66996/archive/master.zip)





### Source Code

Step 1: <b>Initialize and Get the Image</b></p>
<p><code>Dim Photo as StdPicture<br>
Set Photo = LoadPicture(&quot;Photo.jpg&quot;)</code></p>
<p>Step 2: <b>Set Output Size</b></p>
<p><code>MaxHeight = 1400&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;'1.00 inch <br>
MaxWidth = 1050&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;'0.75 inch </code></p>
<p>Step 3: <b>Determine the Scale Factor</b> </p>
<p><code>'The dimensions of the image file <br>
	picHeight = Photo.Height<br>
picWidth = Photo.Width<br>
<br>
If picHeight &gt; picWidth Then<br>
&nbsp;&nbsp;'if height is more than width<br>
 &nbsp;&nbsp;'get Scale factor based on height<br>
</code><code>&nbsp;&nbsp;sFactor = MaxHeight / picHeight&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<br>
Else<br>
&nbsp;&nbsp;'if height is less than width<br>
&nbsp;&nbsp;'get Scale factor based on height<br>
</code><code>&nbsp;&nbsp;sFactor = MaxWidth / picWidth<br>
End If</code></p>
<p>Step 4: <b>The final paint command</b> </p>
<p><code>printer.PaintPicture Photo, StartLeft, StartTop, picWidth * sFactor, picHeight * sFactor <br>
</code></p>
<p>Replace the <code>printer</code> with the name of the form or picture box are required if you want to display the image on the screen.</p><p>Here's the complete code.</p>
<p><code>Dim Photo as StdPicture<br>
Set Photo = LoadPicture(&quot;Photo.jpg&quot;)<br>
</code><code>picHeight = Photo.Height<br>
picWidth = Photo.Width<br>
If picHeight &gt; picWidth Then<br>
&nbsp;&nbsp;'if height is more than width<br>
&nbsp;&nbsp;'get Scale factor based on height<br>
</code><code>&nbsp;&nbsp;sFactor = MaxHeight / picHeight&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<br>
Else<br>
&nbsp;&nbsp;'if height is less than width<br>
&nbsp;&nbsp;'get Scale factor based on height<br>
</code><code>&nbsp;&nbsp;sFactor = MaxWidth / picWidth<br>
End If<br>
Printer.PaintPicture Photo, StartLeft, StartTop, picWidth * sFactor, picHeight * sFactor</code></p>

