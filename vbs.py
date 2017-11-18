import subprocess

managerText = """
Sub includeFile(fSpec)
    With CreateObject("Scripting.FileSystemObject")
       executeGlobal .openTextFile(fSpec).readAll()
    End With
End Sub


"""

class VBS:

	def __init__ (self, file):
		global managerText
		self.file = file
		f = open("__manager.vbs", "w+")
		managerText += 'includeFile("' + file + '") \nEval(WScript.Arguments.Item(0))'
		f.write(managerText)
		f.close()
		
	def run(self, code):
		result = subprocess.run(["cmd", "/C", "cscript", '__manager.vbs ', code], stdout=subprocess.PIPE, stderr=subprocess.PIPE)
		if result.stderr is '':
			raise Exception("Error occured")
		else:
			return self.parse(result.stdout)
	
	def parse(self, str):
		str = str.decode()
		str = str.strip()
		n = str.index("reserved") + len("reserved") + 1
		str = str[n:]
		str = str.strip()
		return str

