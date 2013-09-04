" File: win32htmlfmt.vim
" Author: wz520 [wingzero1040@gmail.com]
" Baidu Tieba ID: 天使的枷锁
" Last Modified: 2013-09-04
"
" 本插件提供一系列函数用于从 Windows 剪贴板中粘贴 HTML Format 格式的数据到 Vim
"   中，即：可以粘贴网页浏览器中复制的内容的 HTML 源代码，方便编写 HTML。
"   测试坏境： WinXP SP3 + IE8, Firefox 23, Chrome 29
"              Win8 64bit + IE10, Chrome 29
"
" 关于 HTML Format 的详细信息请参阅 MSDN：
" http://msdn.microsoft.com/en-us/library/aa767917%28v=vs.85%29.aspx
" http://msdn.microsoft.com/en-us/library/windows/desktop/ms649015%28v=vs.85%29.aspx
"
" Note:
" 在 GTK2 环境下 Vim 自带类似的功能，无需任何插件，并支持复制。
"   只需 set clipboard+=html 即可。详见 :help 'clipboard' 。
"
" Usage:
" 【注意】：使用本插件的任何函数前必须先设置 encoding 为 utf-8 ！
"   :set encoding=utf-8
"
" 需配合 win32htmlfmt.dll 使用。所以只支持 Windows 。
" 而且因为 dll 是 32 位的，所以只支持 32 位的 Vim/GVim（不过貌似运行在 64 位的
" Windows 下没问题，只要 Vim/GVim 也是 32 位的）。
" 如果想要 64 位的 DLL，可以尝试拿压缩包里的 win32htmlfmt.c 来编译（如果不需要
" 自己编译，那 .c 文件就没用了，任君处置 :）。
"
" 解压到 $VIM/vimfiles/ , 然后启动 Vim 后在任何可以执行 Ex 命令的地方执行:
"	:call win32htmlfmt#pasteFragment([opt])
"		将剪贴板内 HTML Format 格式的 Fragment 数据插入（粘贴）到当前行下面。
"		若要插到当前行上面，只需将包含 'above' 的字符串作为参数传入即可，例：
"		  :call win32htmlfmt#pasteFragment('above')
"		关于 Fragment 的详细信息请参阅上面的 URL。
"		简而言之就是只有选区内容，不包括任何其他杂七杂八的东西。
"	:call win32htmlfmt#getFragment()
"		与 pasteFragment() 相似，区别在于将原本用于粘贴的内容作为字符串返回。
"		如果需要转换成列表，推荐 win32htmlfmt#toList() 。
"	:call win32htmlfmt#pasteAll([opt])
"		将剪贴板内 HTML Format 格式的所有内容插入到当前行下面。
"		这会包括开头的 Description 部分，以及其他非选区部分的东西。
"		其他同 pasteFragment()
"	:call win32htmlfmt#getAll()
"		与 pasteAll() 相似，区别在于将原本用于粘贴的内容作为字符串返回。
"		如果需要转换成列表，推荐使用下面的 win32htmlfmt#toList() 。
"	:call win32htmlfmt#toList(s)
"	    将字符串 s 以 \n 分割成列表，并删除行尾的 \r （在 Vim 里通常显示为 ^M）
" 
" 2013-09-04 版新增以下函数：
"   :call win32htmlfmt#pasteKeyword(kwd)
"		获取 Description 中指定 Keyword 的值，并粘贴到当前行下面。
"		kwd 参数指定要获取的 Keyword 。
"		[opt] 参数的用法请参阅上面的 pasteFragment() 或 pasteAll() 函数。
"		比较有用的 Keyword 是 "SourceURL"，可以获取被复制页面的 URL 地址。
"		其他 Keyword 请参阅 MSDN 。
"   :call win32htmlfmt#getKeyword(kwd)
"		与 pasteKeyword() 相似，区别在于将原本用于粘贴的内容作为字符串返回。
"
" 
" paste* 系列函数在当剪贴板中没有 HTML Format 格式的数据时会显示一条提示信息，
" 而 get* 系列函数不会（返回空串表示没有数据）。
" 
" 当然每次都敲命令有点蛋疼，你可以自己 map 到喜欢的键上，例如在 _vimrc 中加入：
"  :nmap <F4> :call win32htmlfmt#pasteFragment()<CR>
"  :nmap <S-F4> :call win32htmlfmt#pasteFragment('above')<CR>

if v:version < 700
	finish
endif

let s:cpo_save = &cpo
set cpo&vim

" Init
let s:dll_path = expand("<sfile>:r")

let s:opt_nodatamsg = 0  " Echo message when no data in the clipboard

" Convert "HTML Format" data to a dict
"
" the dict returned by this function contains all Description keywords:
" e.g. dict['StartHTML'] contains the value of StartHTML keyword of Description
"
" In addition, it contains an extra key called "all" which contains the entire
" "HTML Format" data.
func s:ToDict(htmlformat)
	let input = a:htmlformat
	let output = {}

	" Get keywords.
	" Only first 1024 bytes is used in order to deal with too big html
	" We need to know the end position of the keywords, which is the value of
	" "StartHTML" keyword.
	let leadinglines = input[0:1023]
	let endpos = matchstr(leadinglines, '\nStartHTML:\zs\d\+\ze')
	if empty(endpos)
		return {}
	else
		let endpos = str2nr(endpos, 10)
		let output['all'] = input
	endif

	let pos = 0
	let leadinglineslist = split(leadinglines, "\n")
	unlet leadinglines " avoid using it accidentally

	" -- Get each line, until endpos is reached or no more keywords
	for line in leadinglineslist
		let keyvalue = split(line, ":")
		if len(keyvalue) < 2
			break  " no more keywords
		endif

		let key = keyvalue[0]
		let value = join(keyvalue[1:], ':')
		let value = substitute(value, '\r$', '', '') " remove trailing CR
		let output[key] = value

		" reached endpos?
		let pos += len(line) + 1
		if pos >= endpos
			break
		endif
	endfor

	return output
endfunc

func s:GetHTMLFormat()
	let dict = {}

	let opt_nodatamsg = s:opt_nodatamsg
	let s:opt_nodatamsg = 0  " Reset to default

	if has('win32') && has('libcall') && &encoding == 'utf-8'
		let result = libcall(s:dll_path, 'GetHTMLFormat', 0)
		if !empty(result)
			let dict = s:ToDict(result)
		elseif opt_nodatamsg
			echo 'No "HTML Format" data in the clipboard'
		endif
	else
		echohl ErrorMsg
		echo "win32htmlfmt: cannot use this function because one or more of:\n"
			\ 	"* You are using a non-win32 version of Vim or GVim;\n"
			\ 	"* +libcall feature is not compiled in;\n"
			\ 	"* The value of 'encoding' option is not 'utf-8', try :set enc=utf-8\n"
		echohl None
	endif

	return dict
endfunc

" If one of the parameters is "", return dict['all']
func win32htmlfmt#getRange(start_kwd, end_kwd)
	let dict = s:GetHTMLFormat()
	if empty(dict)
		return ""
	elseif empty(a:start_kwd) || empty(a:end_kwd)
		return dict['all']
	else
		let spos = str2nr(dict[(a:start_kwd)], 10)
		let epos = str2nr(dict[(a:end_kwd)], 10)
		return dict['all'][(spos):(epos-1)]
	endif
endfunc

" Get the value of the keyword of "HTML Format" data
func win32htmlfmt#getKeyword(kwd)
	let dict = s:GetHTMLFormat()
	return empty(dict) ? "" : dict[(a:kwd)]
endfunc

func win32htmlfmt#pasteKeyword(kwd, ...)
	let s:opt_nodatamsg = 1  " Show message if no data found
	let value = win32htmlfmt#getKeyword(a:kwd)
	if empty(value)
		return
	endif

	let inspoint = line('.')
	if len(a:000) > 0 && a:000[0] =~# 'above'
		let inspoint -= 1
	endif
	call append(inspoint, value)
endfunc

func win32htmlfmt#toList(s)
	let slist = split(a:s, "\n")
	call map(slist, 'substitute(v:val, ''\r$'', '''', '''')') " trim ^M
	return slist
endfunc

func win32htmlfmt#pasteRange(start_kwd, end_kwd, opt)
	let s:opt_nodatamsg = 1  " Show message if no data found
	let selection = win32htmlfmt#getRange(a:start_kwd, a:end_kwd)
	if empty(selection)
		return
	endif

	let selectionlist = win32htmlfmt#toList(selection)

	let inspoint = line('.')
	if len(a:opt) > 0 && a:opt[0] =~# 'above'
		let inspoint -= 1
	endif
	call append(inspoint, selectionlist)
endfunc

func win32htmlfmt#getFragment()
	return win32htmlfmt#getRange("StartFragment", "EndFragment")
endfunc

func win32htmlfmt#pasteFragment(...)
	call win32htmlfmt#pasteRange("StartFragment", "EndFragment", a:000)
endfunc

func win32htmlfmt#getAll()
	return win32htmlfmt#getRange("", "")
endfunc

func win32htmlfmt#pasteAll(...)
	call win32htmlfmt#pasteRange("", "", a:000)
endfunc

" restore 'cpo'
let &cpo = s:cpo_save
unlet s:cpo_save

" vim:ts=4
