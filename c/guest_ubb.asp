<script language="javascript" src="c/guest_ubb.js"></script>
<table border="0" cellspacing="0" cellpadding="2">
  <tr>
    <td><select name="font" onFocus="this.selectedIndex=0" onChange="chfont(this.options[this.selectedIndex].value)" size="1">
      <option value="" selected>选择字体</option>
      <option value="宋体">宋体</option>
      <option value="黑体">黑体</option>
      <option value="Arial">Arial</option>
      <option value="Book Antiqua">Book Antiqua</option>
      <option value="Century Gothic">Century Gothic</option>
      <option value="Courier New">Courier New</option>
      <option value="Georgia">Georgia</option>
      <option value="Impact">Impact</option>
      <option value="Tahoma">Tahoma</option>
      <option value="Times New Roman">Times New Roman</option>
      <option value="Verdana">Verdana</option>
    </select></td>
    <td><select name="size" onFocus="this.selectedIndex=0" onChange="chsize(this.options[this.selectedIndex].value)" size="1">
      <option value="" selected>字体大小</option>
      <option value="1">1</option>
      <option value="2">2</option>
      <option value="3">3</option>
      <option value="4">4</option>
      <option value="5">5</option>
      <option value="6">6</option>
    </select></td>
    <td><select name="color"  onFocus="this.selectedIndex=0" onChange="chcolor(this.options[this.selectedIndex].value)" size="1">
      <option value="" selected>字体颜色</option>
      <option value="Black" style="background-color:black;color:black;">Black</option>
      <option value="White" style="background-color:white;color:white;">White</option>
      <option value="Red" style="background-color:red;color:red;">Red</option>
      <option value="Yellow" style="background-color:yellow;color:yellow;">Yellow</option>
      <option value="Pink" style="background-color:pink;color:pink;">Pink</option>
      <option value="Green" style="background-color:green;color:green;">Green</option>
      <option value="Orange" style="background-color:orange;color:orange;">Orange</option>
      <option value="Purple" style="background-color:purple;color:purple;">Purple</option>
      <option value="Blue" style="background-color:blue;color:blue;">Blue</option>
      <option value="Beige" style="background-color:beige;color:beige;">Beige</option>
      <option value="Brown" style="background-color:brown;color:brown;">Brown</option>
      <option value="Teal" style="background-color:teal;color:teal;">Teal</option>
      <option value="Navy" style="background-color:navy;color:navy;">Navy</option>
      <option value="Maroon" style="background-color:maroon;color:maroon;">Maroon</option>
      <option value="LimeGreen" style="background-color:limegreen;color:limegreen;">LimeGreen</option>
    </select></td>
    <td><img src="images/bb_bold.gif" border="0" alt="粗体" onClick="ubbFormat('B')"></td>
    <td><img src="images/bb_italicize.gif" alt="斜体" width="23" height="22" border="0" onClick="ubbFormat('I')"></td>
    <td><img src="images/bb_underline.gif" border="0" alt="下划线" onClick="ubbFormat('U')"></td>
    <td><img src="images/bb_center.gif" border="0" alt="居中对齐" onClick="ubbFormat('CENTER')"></td>
    <td><img src="images/bb_email.gif" border="0" alt="插入EMAIL地址" onClick="ubbFormat('EMAIL')"></td>
    <td><img src="images/bb_url.gif" border="0" alt="插入网址" onClick="ubbFormat('URL')"></td>
    <td><img src="images/bb_image.gif" border="0" alt="插入图片" onClick="ubbInsert('IMG')"></td>
  </tr>
</table>