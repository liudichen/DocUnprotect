# DocUnprotect for windows-Word文档解密（解除保护）工具
>  经常地，某些同事的Word文档经常由于各种原因导致文档被保护，即需要输入密码解除保护后才能修改，而密码又是未知的。
>  几年前，xxc同学写了一个VB版本的，当时只能解除.doc扩展名的文档，执行速度很快。但近日使用时，经常即使是.doc格式（另存的，不是直接改扩展名），也无法解除保护。
####  基于上述原因，本人准备自己尝试搞了一个试下（python学了一点点，纯新手）：
1.  使用python写的，所以生成的exe文件会体积巨大，这个是语言劣势，无法消除。所以应该优先使用之前版本的试下，不行再用这个再抢救下看看。执行速度显著的慢。
2.  基于.docx格式采用的方法，所以实际会先把.wps（必须安装了WPS才行）、.doc、.docx格式先尝试另存为.docx格式，这个过程不快。除这3种格式外，其他格式并不会执行处理。
3.  会调用本机Word或WPS，所以需要本机安装有至少一种，如果是.wps格式，则必须安装WPS。
4.  解除保护后会在文件所在目录生成一个“解除保护de-”前缀的.docx格式的文档。
5.  理论上可以同时拖拽多个文档进行解除保护（不过老夫没试）。
6.  电脑装有DGS等加密软件可能会导致未知的失败。
7.  可能存在一些BUG，仅在我的2台电脑上试过。


####  可直接下载右侧release中exe文件使用，或者自行下载使用pyinstaller打包，注意修改.spec中的相关信息（比如路径）。 在工作目录下  pyinstall docxunprotect.spec即可。
