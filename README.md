# convert2pdf

<!-- TOC -->

- [convert2pdf](#convert2pdf)
- [종속성](#종속성)
- [설치](#설치)
- [사용](#사용)

<!-- /TOC -->

워드파일, 파워포인트파일을 pdf로 컨버팅해주는 스크립트 입니다. <br>
`ls -Recurse` 와 연결해서 구석구석 모든 파일을 pdf로 변환하면 편합니다. <br>
pdf문서로 변환해서 태블릿, 전자책리더에서 편하게 문서를 봅시다. ^^ <br>

<br>
<br>
<br>

# 종속성
- MS Word
- MS PowerPoint

<br>
<br>
<br>

# 설치

```powershell
PS> Install-Module -Name convert2pdf
```

<br>
<br>
<br>


# 사용

```powershell
PS> ls -Recurse | convert2pdf

C:\temp\doc\aaa.pptx start ...
C:\temp\doc\aaa.pdf done ...
0
C:\temp\doc\aaa.pdf
C:\temp\doc\bbb.docx start ...
C:\temp\doc\bbb.pdf done ...
0
C:\temp\doc\bbb.pdf
C:\temp\doc\ccc.docx start ...
C:\temp\doc\ccc.pdf done ...
0
C:\temp\doc\ccc.pdf
...
```
<br>
<br>
<br>