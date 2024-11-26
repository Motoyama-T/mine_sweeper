''初期設定'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
option explicit
Randomize
dim test,x,r,i,n,m,p,q,draw,message,hairetu(7,7),tmpHairetu,abcHairetu,sen1,sen2,blank,hide,bomb,block,fumei
test = "テスト"
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

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''メイン処理'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
SetBomb()
DO until isEmpty(x)
MasDraw()
x = inputbox("クリアまであと " &tmpHairetu &" マス" &vbCrLf &draw,"マインスイーパー",message)
if x = "reset" then
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
   hairetu(i,n)(2) = 1
   draw = draw &blank
  elseif hairetu(i,n)(2) = 0 then
   draw = draw &hide
  elseif hairetu(i,n)(2) = 2 then
   draw = draw &block
  elseif hairetu(i,n)(2) = 3 then
   draw = draw &fumei
  elseif hairetu(i,n)(1) = "bomb" then
   draw = draw &bomb
  else
   draw = draw &"|　"&(hairetu(i,n)(1)) &" "
  end if
  if hairetu(i,n)(2) = 0 then
   tmpHairetu = tmpHairetu + 1
  end if
 next
 draw = draw &"|" &vbCrLF
next
draw = draw &sen2
end function
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''爆弾を設置する関数'''''''''''''''''''''''''''''''''''''''''''''''''
function SetBomb()
for i = 0 to r step 1
 for n = 0 to r step 1
  hairetu(i,n) = array(i+1&abcHairetu(n),0,0)
 next
next
m = 0
do until m >= 10
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
end function
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''マスを開く関数'''''''''''''''''''''''''''''''''''''''''''''''''''''
function MasOpen()
for i = 0 to r step 1
 for n = 0 to r step 1
 if x = hairetu(i,n)(0) then
  hairetu(i,n)(2) = 1
 elseif x = "-" &(hairetu(i,n)(0)) then
  hairetu(i,n)(2) = 0
 elseif x = "+" &(hairetu(i,n)(0)) then
  hairetu(i,n)(2) = 2
 elseif x = "?" &(hairetu(i,n)(0)) then
  hairetu(i,n)(2) = 3
 end if
 next
next
end function
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''