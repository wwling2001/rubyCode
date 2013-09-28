
#coding:utf-8
####encoding:utf-8
#
#测试读取log的路径
#http://bluefish-win.googlecode.com/svn/trunk/
#http://yet-another-music-application.googlecode.com/svn/trunk


require 'win32ole'
#require 'Dir'
require 'pathname'

#这里直接通过块注释，先将有问题的代码注释掉

class SvnOper
    attr_reader:cmdStr, :httpPath, :userName, :password

    def initialize(http, userName, passWord)
        @httpPath = http
        @userName = userName
        @password = passWord
    end

    def svnLogCmd()
        if (@userName == "") or (@password == "")
            @cmdStr = "svn log -v " + @httpPath
        else
            @cmdStr = "svn log -v " + @httpPath + " --username " + @userName + " --password " + @password
        end
        
    end

    def excCmd()
	filename = ".\\svnlog.txt"
        svnLogCmd()
        outPut = `#{cmdStr}`
        puts "cmd excute......"
	file=File.new(filename, "w")
	file.print(outPut)
	file.close
	puts "write file successful!"
        #puts outPut
    end

    def parseLogInfo()
    	filename = ".\\svnlog2.txt"
	
	#将文件读成一个字符串数组
	filearr = IO.readlines(filename)
	
	#将文件读到一个字符串变量中
	filecontent = IO.read(filename)
	#file = File.new(filename, "r")

	#filearr[0]是svn log输出的记录的分割符，然后用这个分隔符将整个内容切分开来	
	puts filearr[0]
	fileContentArray = filecontent.split(filearr[0])
	
	puts "array size = " 
	puts fileContentArray.size

	n = 1
	while n < fileContentArray.size
		puts "*********************************************"
		puts fileContentArray[n]
		n = n + 1
		puts "*********************************************"
	end
    end
end

=begin
  
class ExcelOp
    attr_reader:exApp, :exdbook, :exsheet, :exsheets

    def initialize()
        @exApp = WIN32OLE.new('Excel.Application')
        @exApp.visible = true
        @exApp.displayalerts = false
        #@exdbook = 
    end

    def openExcelFile(fileName)
          #判断文件是否存在，如果不存在侧创建并保存
          if File.Exist(fileName) == false
              @exdbook = @exApp.Workbooks.add()
              @exsheets = @exdbook.Worksheets
              nCnt = 1
              while nCnt <= @exsheets.Count()
                  @exsheets(nCnt).Name = "LogOut" + nCnt
              end
              @exdbook.SaveAs(File.dirname(__FILE__) + fileName)
          else
              @exdbook = @exApp.Workbooks.Open(fileName)
          end
    end

    def excelQuit()
          @exdbook.close()
          @exApp.Quit()
    end
end

=end

svnObj = SvnOper.new("http://yet-another-music-application.googlecode.com/svn/trunk", "", "")
#svnObj.excCmd()
svnObj.parseLogInfo()


=begin
class TestCode

    def initialize()
    end

    def doPut()
        puts "Hello world"
        puts File.dirname(__FILE__)
        #puts Dir.pwd
        puts Pathname.new(File.dirname(__FILE__)).realpath

        excel = WIN32OLE.new('Excel.Application')
        excel.visible = true
        excel.displayalerts = false
        workbook = excel.Workbooks.Add()
        worksheet = workbook.Worksheets(1)
        worksheet.Range("A1:D1").value = ["North","South","East","West"];
        worksheet.Range("A2:B2").value = [5.2, 10]
        worksheet.Range("C2").value = 8
        worksheet.Range("D2").value = 20
        worksheet.Range("F3").value = "中国"

        range = worksheet.Range("A1:D2")
        range.select
        chart = workbook.Charts.Add

        workbook.SaveAs("E:\\Programming\\Script\\Ruby\\myfiletest.xls")
        workbook.saved = true

        workbook.close()

        #excel.ActiveWorkbook.Close(0)
        excel.Quit()
    end
end

justPut = TestCode.new
justPut.doPut()

=end

=begin
puts File.dirname(__FILE__)
#puts Dir.pwd
puts Pathname.new(File.dirname(__FILE__)).realpath

excel = WIN32OLE.new('Excel.Application')
excel.visible = true
excel.displayalerts = false
workbook = excel.Workbooks.Add()
worksheet = workbook.Worksheets(1)
worksheet.Range("A1:D1").value = ["North","South","East","West"];
worksheet.Range("A2:B2").value = [5.2, 10]
worksheet.Range("C2").value = 8
worksheet.Range("D2").value = 20
worksheet.Range("F3").value = "中国"

range = worksheet.Range("A1:D2")
range.select
chart = workbook.Charts.Add

workbook.SaveAs("E:\\Programming\\Script\\Ruby\\myfiletest.xls")
workbook.saved = true

workbook.close()

#excel.ActiveWorkbook.Close(0)
excel.Quit()

=end

