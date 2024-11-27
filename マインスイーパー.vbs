''操作コマンド'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'■マスを指定する（開く）''''''''''例：1a [行(数字)+列(アルファベット)]
'■マスが開かないようブロックする'''例：+1a
'■ブロックや印を解除する''''''''''例：-1a
'■印をつける'''''''''''''''''''''例：*1a
'■リセットして初めから''''''''''''r
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''初期設定'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
option explicit
Randomize
dim test,mine,x,r,i,n,m,p,q,draw,message,hairetu(7,7),tmpHairetu,abcHairetu,sen1,sen2,blank,hide,bomb,block,fumei,cleartrigger,endtrigger,result,tb,lr

''地雷の個数''''''''''''''''''''''''''''''''''''''''''''''''''''''''

mine =11

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

x = ""
r = 7
abcHairetu = array("a","b","c","d","e","f","g","h")
sen1 = "　　| Ａ | Ｂ | Ｃ | Ｄ | Ｅ | Ｆ | Ｇ | Ｈ |" &vbCrLf
sen2 = "ーー+ーー+ーー+ーー+ーー+ーー+ーー+ーー+ーー+" &vbCrLf
blank = "| 　 "
hide = "|////"
bomb = "| ★ "
block = "|/☆/"
fumei = "|/？/"
cleartrigger = false
endtrigger = false
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''メイン処理'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
SetBomb()
DO until isEmpty(x)
MasDraw()
Ending()
x = inputbox(message,"マインスイーパー")
if x = "r" then
 SetBomb()
else
 MasOpen()
end if
LOOP
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''描画する関数'''''''''''''''''''''''''''''''''''''''''''''''''''''''
function MasDraw()
tmpHairetu = 0
draw = sen1
for i = 0 to r step 1
 draw = draw &sen2
 draw = draw &" " &string(2-len(i),"0") &i+1 &" "
 for n = 0 to r step 1
  if hairetu(i,n)(1) = "0" then
   draw = draw &blank
  elseif hairetu(i,n)(2) = 0 then
   draw = draw &hide
  elseif hairetu(i,n)(2) = 2 then
   draw = draw &block
  elseif hairetu(i,n)(2) = 3 then
   draw = draw &fumei
  elseif hairetu(i,n)(1) = "bomb" then
   draw = draw &bomb
   endtrigger = true
  else
   draw = draw &"|　"&(hairetu(i,n)(1)) &" "
  end if
  if hairetu(i,n)(2) <> 1 then
   tmpHairetu = tmpHairetu + 1
  end if
 next
 draw = draw &"|" &vbCrLF
next
draw = draw &sen2
if tmpHairetu - mine = 0 then
 cleartrigger = true
end if
end function
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''爆弾を設置する関数'''''''''''''''''''''''''''''''''''''''''''''''''
function SetBomb()
cleartrigger =false
endtrigger = false
for i = 0 to r step 1
 for n = 0 to r step 1
  hairetu(i,n) = array(i+1&abcHairetu(n),0,0)
 next
next
m = 0
do until m >= mine
 i = int(8*rnd)
 n = int(8*rnd)
 if hairetu(i,n)(1) = "0" then
  for p = i-1 to i+1 step 1
   for q = n-1 to n+1 step 1
    if p>=0 and q>=0 and p<=r and q<=r then
     hairetu(p,q)(1) = (hairetu(p,q)(1))+1
    end if
   next
  next
  hairetu(i,n)(1) = "bomb"
  m = m + 1
 end if
loop
for p = 0 to r
 for q = 0 to r
  if hairetu(p,q)(1) = "0" then
   for tb = -1 to 1
    for lr = -1 to 1
     if p + tb >= 0 and p + tb <= r and q + lr >= 0 and q + lr <= r then
      hairetu(p + tb,q + lr)(2) = 1
     end if
    next
   next
  end if
 next
next
end function
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''マスを開く関数'''''''''''''''''''''''''''''''''''''''''''''''''''''
function MasOpen()
for i = 0 to r step 1
 for n = 0 to r step 1
 if x = hairetu(i,n)(0) then
  if hairetu(i,n)(2) <> 2 then
   hairetu(i,n)(2) = 1
  end if
 elseif x = "-" &(hairetu(i,n)(0)) then
  hairetu(i,n)(2) = 0
 elseif x = "+" &(hairetu(i,n)(0)) then
  hairetu(i,n)(2) = 2
 elseif x = "*" &(hairetu(i,n)(0)) then
  hairetu(i,n)(2) = 3
 end if
 next
next
end function
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''地雷を踏んだ関数'''''''''''''''''''''''''''''''''''''''''''''''''''
sub Ending()
if endtrigger then
 for p = 0 to r
  for q = 0 to r
   hairetu(p,q)(2) = 1
  next
 next
 result = tmpHairetu + 1 - mine
 MasDraw()
 message = "！！！！爆発した！！！！" &vbCrLf &"クリアまであと " &result &" マスでした" &vbCrLf &draw
else
 result = tmpHairetu - mine
 if cleartrigger then
  message = "★★★★★★★★★★★★★★★★★★★★★★★" &vbCrLf &"∩(＾ω＾∩)★★★！クリア！★★★(∩・ω・)∩" &vbCrLf &"★★★★★★★★★★★★★★★★★★★★★★★" &vbCrLf &draw
 else
  message = "クリアまであと " &result &" マス" &vbCrLf &draw
 end if
end if
end sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
