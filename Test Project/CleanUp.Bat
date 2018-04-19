@Echo Off

attrib -s -h -r

Del *.Ncb
Del *.Opt
Del *.Plg
Del *.Aps
Del *.Scc
Del *.suo
Del *.xml
Del *.old
Del *.user
Del *.sdf

Del Release\*.* /Q
Del Debug\*.* /Q

RD Release
RD Debug

Del Resources\*.aps

Cls