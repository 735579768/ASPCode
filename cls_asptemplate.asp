<%
'    ASP Template 1.2.1
'    Copyright (C) 2001-2004 Valerio Santinelli
'
'    This library is free software; you can redistribute it and/or
'    modify it under the terms of the GNU Lesser General Public
'    License as published by the Free Software Foundation; either
'    version 2.1 of the License, or (at your option) any later version.
'
'    This library is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
'    Lesser General Public License for more details.
'
'    You should have received a copy of the GNU Lesser General Public
'    License along with this library; if not, write to the Free Software
'    Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'
'   ---------------------------------------------------------------------------
'
' ASP Template main class file
'
' Author: Valerio Santinelli <tanis@altralogica.it>
' $Id: asptemplate.asp 52 2007-01-26 06:54:41Z mrtwice $
'

'===============================================================================
' Name: ASPTemplate Class
' Purpose: HTML separation class
' Functions:
'     <functions' list in alphabetical order>
' Properties:
'     <properties' list in alphabetical order>
' Methods:
'     <Methods' list in alphabetical order>
' Author: Valerio Santinelli <tanis@mediacom.it>
' Start: 2001/01/01
' Modified: 2001/12/19
'===============================================================================
class ASPTemplate

	' Contains the error objects
	private p_error
	
	' Print error messages?
	private p_print_errors
	
	' What to do with unknown tags (keep, remove or comment)?
	private p_unknowns
	
	' Opening delimiter (usually "{{")
	private p_var_tag_o
	
	' Closing delimiter (usually "}}")
	private p_var_tag_c

	'private p_start_block_delimiter_o
	'private p_start_block_delimiter_c
	'private p_end_block_delimiter_o
	'private p_end_block_delimiter_c
	
	'private p_int_block_delimiter
	
	private p_template
	private p_variables_list
	private p_blocks_list
	private p_blocks_name_list
	private	p_regexp
	private p_parsed_blocks_list

	private p_boolSubMatchesAllowed
	
	' Directory containing HTML templates
	public p_templates_dir
	
	'===============================================================================
	' Name: class_Initialize
	' Purpose: Constructor
	' Remarks: None
	'===============================================================================
	private sub class_Initialize
		p_print_errors = FALSE
		p_unknowns = "keep"
		' Remember that opening and closing tags are being used in regular expressions
		' and must be explicitly escaped
		p_var_tag_o = "\{\{"
		p_var_tag_c = "\}\}"
		' Block delimiters are actually disabled and no longer available. Maybe they'll be again
		' in the future.
		'p_start_block_delimiter_o = "<!-- BEGIN "
		'p_start_block_delimiter_c = " -->"
		'p_end_block_delimiter_o = "<!-- END "
		'p_end_block_delimiter_c = " -->"
		'p_int_block_delimiter = "__"
		p_templates_dir = "my_templates/"
		set p_variables_list = createobject("Scripting.Dictionary")
		set p_blocks_list = createobject("Scripting.Dictionary")
		set p_blocks_name_list = createobject("Scripting.Dictionary")
		set p_parsed_blocks_list = createobject("Scripting.Dictionary")
		p_template = ""
		p_boolSubMatchesAllowed = not (ScriptEngineMajorVersion & "." & ScriptEngineMinorVersion & "." & ScriptEngineBuildVersion < "5.5")
		Set p_regexp = New RegExp   
	end sub
	
	'===============================================================================
	' Name: SetTemplatesDir
	' Input:
	'    dir as Variant Directory
	' Output:
	' Purpose: Sets the directory containing html templates
	' Remarks: None
	'===============================================================================
	public sub SetTemplatesDir(dir)
		p_templates_dir = dir
	end sub

	'===============================================================================
	' Name: SetTemplate
	' Input:
	'    template as Variant String containing the template
	' Output:
	' Purpose: Sets a template passed through a string argument
	' Remarks: None
	'===============================================================================
	public sub SetTemplate(template)
		p_template = template
	end sub
	
	'===============================================================================
	' Name: GetTemplate
	' Input:
	' Output:
	'    template as Variant String
	' Purpose: returns template as a string
	' Remarks: None
	'===============================================================================
	public function GetTemplate
		GetTemplate = p_template
	end function
	
	'===============================================================================
	' Name: SetUnknowns
	' Input:
	'    action as String containing the action to perform with unrecognized
	'    tags in the template
	' Output:
	' Purpose: Sets a variable passed through a string argument
	' Remarks: The action can be one of the following:
	'  - 'keep': leave the tags untouched
	'  - 'remove': remove the tags from the output
	'  - 'comment': mark the tags as HTML comment
	'===============================================================================
	public sub SetUnknowns(action)
		if (action <> "keep") and (action <> "remove") and (action <> "comment") then
			p_unknowns = "keep"
		else
			p_unknowns = action
		end if
	end sub

	'===============================================================================
	' Name: SetTemplateFile
	' Input:
	'    inFileName as Variant Name of the file to read the template from
	' Output:
	' Purpose: Sets a template given the filename to load the template from
	' Remarks: None
	'===============================================================================
	public sub SetTemplateFile(inFileName)
		dim FSO, oFile
		if len(inFileName) > 0 then
			set FSO = createobject("Scripting.FileSystemObject")
			'response.write server.mappath(p_templates_dir & inFileName)
			if FSO.FileExists(server.mappath(p_templates_dir & inFileName)) then
				set oFile = FSO.OpenTextFile(server.mappath(p_templates_dir & inFileName), 1)
				p_template = oFile.ReadAll
				oFile.Close
				set oFile = nothing
			else
				response.write "<b>ASPTemplate Error: File [" & inFileName & "] does not exists!</b><br>"
			end if
			set FSO = nothing
		else
			response.write "<b>ASPTemplate Error: SetTemplateFile missing filename.</b><br>"
		end if
		
	end sub
	

	'===============================================================================
	' Name: SetVariable
	' Input:
	'    s as Variant - Variable name
	'    v as Variant - Value
	' Output:
	' Purpose: Sets a variable given it's name and value
	' Remarks: None
	'===============================================================================
	public sub SetVariable(s, v)
		if p_variables_list.Exists(s) then
			p_variables_list.Remove s
			p_variables_list.Add s, v
		else
			p_variables_list.Add s, v
		end if
	end sub


	'===============================================================================
	' Name: Append
	' Input:
	'    s as Variant - Variable name
	'    v as Variant - Value
	' Output:
	' Purpose: Sets a variable appending the new value to the existing one
	' Remarks: None
	'===============================================================================
	public sub Append(s, v)
		Dim tmp
		if p_variables_list.Exists(s) then
			tmp = p_variables_list.Item(s) & v
			p_variables_list.Remove s
			p_variables_list.Add s, tmp
		else
			p_variables_list.Add s, v
		end if
	end sub
	
	
	'===============================================================================
	' Name: SetVariableFile
	' Input:
	'    s as Variant Variable name
	'    inFileName as Variant Name of the file to read the value from
	' Output:
	' Purpose: Load a file into a variable's value
	' Remarks: None
	'===============================================================================
	public sub SetVariableFile(s, inFileName)
		if len(inFileName) > 0 then
			dim FSO, oFile
			set FSO = createobject("Scripting.FileSystemObject")
			if FSO.FileExists(server.mappath(p_templates_dir & inFileName)) then
				set oFile = FSO.OpenTextFile(server.mappath(p_templates_dir & inFileName))
				ReplaceBlock s, oFile.ReadAll
				oFile.Close
				set oFile = Nothing
			else
				response.write "<b>ASPTemplate Error: File [" & inFileName & "] does not exists!</b><br>"
			end if
			set FSO = nothing
		else
			'Filename was never passed!
		end if
	end sub


	'===============================================================================
	' Name: ReplaceBlock
	' Input:
	'    s as Variant Variable name
	'    inFile as Variant Content of the file to place in the template
	' Output:
	' Purpose: Function used by SetVariableFile to load a file and replace it
	'          into the template in place of a variable
	' Remarks: None
	'===============================================================================
	public sub ReplaceBlock(s, inFile)
		p_regexp.IgnoreCase = True
		p_regexp.Global = True
		SetVariable s, inFile
		p_regexp.Pattern = p_var_tag_o & s & p_var_tag_c
		p_template = p_regexp.Replace(p_template, inFile)   
	end sub

	public property get GetOutput
	on error resume next
		Dim Matches, match, MatchName
		
		'Replace the variables in the template
		p_regexp.IgnoreCase = True
		p_regexp.Global = True
		p_regexp.Pattern = "(" & p_var_tag_o & ")([^}]+)" & p_var_tag_c
		Set Matches = p_regexp.Execute(p_template)   
		for each match in Matches
			if p_boolSubMatchesAllowed then
				MatchName = match.SubMatches(1)
			else
				MatchName = mid(match.Value,3,Len(match.Value) - 4)
			end if
			if p_variables_list.Exists(MatchName) then
				p_regexp.Pattern = match.Value
				'vardump(p_variables_list.Item(MatchName))
				if not IsObject(p_variables_list.Item(MatchName)) then
				p_template = p_regexp.Replace(p_template, p_variables_list.Item(MatchName))
				else
				p_template = p_regexp.Replace(p_template,"...")
				end if
			end if
			'response.write.write "GetOutput (match): " & match.Value & "<br>"
		next
        
        'this removes any block placeholder tags that are left over
		p_regexp.Pattern = "__[_a-z0-9]*__"
		Set Matches = p_regexp.Execute(p_template)   
		for each match in Matches
			'response.write.write "[[" & match.Value & "]]<br>"
			p_regexp.Pattern = match.Value
			p_template = p_regexp.Replace(p_template, "")
		next

		' deal with unknown tags
		select case p_unknowns
			case "keep"
				'do nothing, leave it
			case "remove"
				'all known matches have been replaced, remove every other match now
				p_regexp.Pattern = "(" & p_var_tag_o & ")([^}]+)" & p_var_tag_c
				Set Matches = p_regexp.Execute(p_template)   
				for each match in Matches
					'response.write.Write "Found match: " & match & "<br>"
					p_regexp.Pattern = match.Value
					p_template = p_regexp.Replace(p_template, "")
				next
			case "comment"
				'all known matches have been replaced, HTML comment every other match
				p_regexp.Pattern = "(" & p_var_tag_o & ")([^}]+)" & p_var_tag_c
				Set Matches = p_regexp.Execute(p_template)   
				for each match in Matches
					p_regexp.Pattern = match.Value
					if p_boolSubMatchesAllowed then
						p_template = p_regexp.Replace(p_template, "<!-- Template variable " & match.Submatches(1) & " undefined -->")
					else
						p_template = p_regexp.Replace(p_template, "<!-- Template variable " & mid(match.Value,3,len(match) - 4) & " undefined -->")
					end if
				next
		end select
		p_template= replace(p_template,"{{void_url}}", "javascript:void(0)")
		p_regexp.IgnoreCase = True
		p_regexp.Global = True
		p_regexp.Pattern = "{{\S*}}"
		p_template = p_regexp.Replace(p_template, "")
		GetOutput = p_template
	end property


	
		
	' TODO: if the block foud contains other blocks, it should recursively update all of them without the needing
	' of doing this by hand.
	public sub UpdateBlock(inBlockName)
		Dim Matches, match, aSubMatch
		Dim braceStart, braceEnd
		
		p_regexp.IgnoreCase = True
		p_regexp.Global = True

		p_regexp.Pattern = "<!--\s+BEGIN\s+(" & inBlockName & ")\s+-->([\s\S.]*)<!--\s+END\s+\1\s+-->"
		Set Matches = p_regexp.Execute(p_template)
		Set match = Matches
		for each match in Matches
			if p_boolSubMatchesAllowed then
				aSubMatch = match.SubMatches(1)
			else
				braceStart = instr(match,"-->") + 3
				braceEnd = instrrev(match,"<!--")
				aSubMatch = mid(match,braceStart,braceEnd - braceStart)
			end if
			'The following check let the user use the same template multiple times
			if p_blocks_list.Exists(inBlockName) then
				p_blocks_list.Remove(inBlockName)
				p_blocks_name_list.Remove(inBlockName)
			end if
			p_blocks_list.Add inBlockName, aSubMatch
			p_blocks_name_list.Add inBlockName, inBlockName
			'printInternalTemplate "UpdateBlock: before replace"
			p_template = p_regexp.Replace(p_template, "__" & inBlockName & "__")
			'printInternalTemplate "UpdateBlock: after replace"
			'response.write.write "[[" & server.HTMLEncode(aSubMatch) & "]]<br>"
		next
	end sub

	public sub ParseBlock(inBlockName)
		Dim Matches, match, tmp, w, aSubMatch
		
		'debugPrint "Parsing: " + inBlockName
		
		w = GetBlock(inBlockName)
		
        'See if there are any sub-blocks in this block
		p_regexp.IgnoreCase = True
		p_regexp.Global = True
		p_regexp.Pattern = "(__)([_a-z0-9]+)__"
		Set Matches = p_regexp.Execute(w)
		Set match = Matches
		for each match in Matches
			if p_boolSubMatchesAllowed then
				aSubMatch = match.SubMatches(1)
			else
				aSubMatch = mid(match.Value,3,len(match) - 4)
			end if

            'If the sub-block has already been parsed, then replace the block
            'identifier with the already parsed text
            if p_parsed_blocks_list.Exists(aSubMatch) then
			    p_regexp.Pattern = "__" & aSubMatch & "__"
				w = p_regexp.Replace(w, p_parsed_blocks_list.Item(aSubMatch))
				p_parsed_blocks_list.Remove(aSubMatch)
			else
			    'if we are here, that means we are parsing a parent block
			    'that has a child block that has not yet been parsed.  We assume
			    'that means that the block should remain empty and we therefore
			    'need to remove the sub-block identifier from the parsed output
			    'of this block.  Otherwise, the sub-block identifiers will be
			    'replaced in the template on future requests to parse the
			    'sub-block.
			    'this removes any block placeholder tags that are left over
    			p_regexp.Pattern = "__" & aSubMatch & "__"
    			w = p_regexp.Replace(w, "")
            end if
		next

		'If this block has already been parsed, append the output to the current
        'entry in the parsed_blocks_list.  Otherwise, create the entry.
		if p_parsed_blocks_list.Exists(inBlockName) then
			tmp = p_parsed_blocks_list.Item(inBlockName) & w
			p_parsed_blocks_list.Remove inBlockName
			p_parsed_blocks_list.Add inBlockName, tmp
		else
			p_parsed_blocks_list.Add inBlockName, w
		end if
        
        'Finally, replace the block identifier in the template with the text of
        'this block and then append the block identifier in case this block
        'is parsed again.
        'If the block is not found in the template, we assume that is because
        'the block is embedded in a parent block that will be parsed in the future.
        'When the parent block is parsed, the content of this block will be included
		p_regexp.IgnoreCase = True
		p_regexp.Global = True
		p_regexp.Pattern = "__" & inBlockName & "__"
		Set Matches = p_regexp.Execute(p_template)
		Set match = Matches
		for each match in Matches
			w = GetParsedBlock(inBlockName)
			p_regexp.Pattern = "__" & inBlockName & "__"
			p_template = p_regexp.Replace(p_template, w & "__" & inBlockName & "__")
		next
        
        'printInternalVariables
        'printInternalTemplate "ParseBlock: end of function ("+inBlockName+")"
	end sub
    
    'Gets the text inside a block, parses and replaces variables, and returns
    'the block of text
	private property get GetBlock(inToken)
		Dim tmp, s
		
		'This routine checks the Dictionary for the text passed to it.
		'If it finds a key in the Dictionary it Display the value to the user.
		'If not, by default it will display the full Token in the HTML source so that you can debug your templates.
		if p_blocks_list.Exists(inToken) then
			tmp = p_blocks_list.Item(inToken)
			s = ParseBlockVars(tmp)
			GetBlock = s
			'response.write.write "s: " & s
		else
			GetBlock = "<!--__" & inToken & "__-->" & VbCrLf
		end if
	end property


	private property get GetParsedBlock(inToken)
		Dim tmp, s
		
		'This routine checks the Dictionary for the text passed to it.
		'If it finds a key in the Dictionary it Display the value to the user.
		'If not, by default it will display the full Token in the HTML source so that you can debug your templates.
		if p_blocks_list.Exists(inToken) then
			tmp = p_parsed_blocks_list.Item(inToken)
			s = ParseBlockVars(tmp)
			GetParsedBlock = s
			'response.write.write "s: " & s
			p_parsed_blocks_list.Remove(inToken)
		else
			GetParsedBlock = "<!--__" & inToken & "__-->" & VbCrLf
		end if
	end property


	public property get ParseBlockVars(inText)
		Dim Matches, match, aSubMatch
		
		p_regexp.IgnoreCase = True
		p_regexp.Global = True

		p_regexp.Pattern = "(" & p_var_tag_o & ")([^}]+)" & p_var_tag_c
		Set Matches = p_regexp.Execute(inText)

		for each match in Matches
			if p_boolSubMatchesAllowed then
				aSubMatch = match.SubMatches(1)
			else
				aSubMatch = mid(match.Value,3,Len(match.Value) - 4)
			end if
			if p_variables_list.Exists(aSubMatch) then
				p_regexp.Pattern = match.Value
				if IsNull(p_variables_list.Item(aSubMatch)) then
					inText = p_regexp.Replace(inText, "")
				else
					inText = p_regexp.Replace(inText, p_variables_list.Item(aSubMatch))
				end if
			end if
			'response.write.write "match.Value: " & match.Value & "<br>"
			'response.write.write "in text: " & inText & "<br>"
		next
		ParseBlockVars = inText
	end property
    
    public sub printInternalVariables()
        'response.write "<b>p_variables_list:</b>"
    	'printr p_variables_list
    	'response.write "<b>p_blocks_list:</b>"
    	'printr p_blocks_list
    	'response.write "<b>p_blocks_name_list:</b>"
    	'printr p_blocks_name_list
    	response.write "<b>p_parsed_blocks_list:</b>"
    	printr p_parsed_blocks_list
    end sub
    
    public sub printInternalTemplate( sPrefix )
        debugPrint sPrefix & "<br /><pre><blockquote>" & Server.HTMLEncode(p_template) & "</blockquote></pre><hr />"
    end sub
    
    private sub debugPrint( sText )
        response.write sText + "<br />"
    end sub
	
	
public sub Parse
Dim parsed
parsed=GetOutput
'parsed = replace(GetOutput,"{{void_url}}", "javascript:void(0)")
'p_regexp.IgnoreCase = True
'p_regexp.Global = True
'p_regexp.Pattern = "{{\S*}}"
'parsed = p_regexp.Replace(parsed, "")
response.write parsed
end sub

public function restr(patrn, replStr)
	parsed = GetOutput
	p_template=Replace(parsed,patrn,replStr)
	end function
public function makehtml(htmlfile)
		dim fs,f
set fs=Server.CreateObject("Scripting.FileSystemObject") 
'response.Write Replace(server.MapPath(htmlfile),"\","/")
'response.End()
on error goto 0
set f=fs.CreateTextFile(Replace(server.MapPath(htmlfile),"\","/"),true)
f.write(GetOutput)
f.close
set f=nothing
set fs=nothing
end function


	public sub ParseReimg
		Dim parsed,srcjs
		srcjs="<script type='text/javascript'>var scrollLoad=(function(options){var defaults=(arguments.length==0)?{src:'xSrc',time:300}:{src:options.src||'xSrc',time:options.time||300};var camelize=function(s){return s.replace(/-(\w)/g,function(strMatch,p1){return p1.toUpperCase();});};this.getStyle=function(element,property){if(arguments.length!=2)return false;var value=element.style[camelize(property)];if(!value){if(document.defaultView&&document.defaultView.getComputedStyle){var css=document.defaultView.getComputedStyle(element,null);value=css?css.getPropertyValue(property):null;}else if(element.currentStyle){value=element.currentStyle[camelize(property)];}}return value=='auto'?'':value;};var _init=function(){var offsetPage=window.pageYOffset?window.pageYOffset:window.document.documentElement.scrollTop,offsetWindow=offsetPage+Number(window.innerHeight?window.innerHeight:document.documentElement.clientHeight),docImg=document.images,_len=docImg.length;if(!_len)return false;for(var i=0;i<_len;i++){var attrSrc=docImg[i].getAttribute(defaults.src),o=docImg[i],tag=o.nodeName.toLowerCase();if(o){postPage=o.getBoundingClientRect().top+window.document.documentElement.scrollTop+window.document.body.scrollTop;postWindow=postPage+Number(this.getStyle(o,'height').replace('px',''));if((postPage>offsetPage&&postPage<offsetWindow)||(postWindow>offsetPage&&postWindow<offsetWindow)){if(tag==='img'&&attrSrc!==null){o.setAttribute('src',attrSrc);}o=null;}}};window.onscroll=function(){setTimeout(function(){_init();},defaults.time);}};return _init();});scrollLoad();</script>"
		
		parsed = GetOutput
		
		response.write replace_img(parsed,"Xsrc","")&srcjs 
	end sub

end class


function replace_img(str,src2,img)
if src2="" then src2="src2"
if img="" then img="/loading.gif"
Dim Matches,temmatchs, a,aa, aSubMatch,zongstr
zongstr=str
Set p_regexp = New RegExp 
p_regexp.IgnoreCase = false
p_regexp.Global = True
p_regexp.Pattern = "<img([^z]*?)/?>"
		Set Matches = p_regexp.Execute(str)	
		for each a in Matches
		p_regexp.Pattern = "src=""(.*?)"""
					set temmatchs = p_regexp.execute(a.value)
					for each aa in temmatchs			
						restr=p_regexp.replace(a.value,replace(aa.value,"src",src2) &" src="""&img&""" ")
						zongstr=replace(zongstr,a.value,restr)
					next
					set temmatchs=nothing
		next
replace_img=zongstr
end function
%>
