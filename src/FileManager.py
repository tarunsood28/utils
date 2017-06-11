from os.path import join
import csv
import sys
import xlrd


class FileSplitter (object):
    
        def __init__(self,file_obj):
            self.file_obj=file_obj

        
        def read_in_chunks (self, chunk_size=1024,sep='\n'):
            incomplete_row = None
            while True:
                chunk = self.file_obj.read(chunk_size)
                if not chunk: # End of file
                    if incomplete_row is not None:
                        yield incomplete_row
                        break
                while True:
                    i = chunk.find(sep)
                    if i == -1:
                        break
                    if incomplete_row is not None:
                        yield incomplete_row + chunk[:i]
                        incomplete_row = None
                    else:
                        yield chunk[:i]
                    chunk = chunk[i+1:]
                if incomplete_row is not None:
                    incomplete_row += chunk
                else:
                    incomplete_row = chunk

    

class DataFileReader (object):
    
    def __init__(self,filepath,filename):
        self.filepath=filepath
        self.filename=filename
        self.file=join(self.filepath,self.filename)

    def delimitedfilereader(self,delimiter):       
        data = {}
        with open (self.file,'r') as f:
            c=0
            reader=csv.DictReader(f, delimiter=delimiter)
            try:
                for row in reader:
                    data[c]=dict(row)
                    c += 1                   
            except csv.Error as e:
                sys.exit('file {}, line {}: {}'.format(self.filename, reader.line_num, e))
        return data
    
    def idelimitedfilereader(self,delimiter):       
        with open (self.file,'r') as f:
            reader=csv.DictReader(f, delimiter=delimiter)
            try:
                for row in reader:
                    yield dict (row)
            except csv.Error as e:
                sys.exit('file {}, line {}: {}'.format(self.filename, reader.line_num, e))
    
    def excelfilereader(self,sheetnames=None):
        data={}
        xl_workbook = xlrd.open_workbook(self.file)
        if sheetnames == None:
            sheetnames = xl_workbook.sheet_names()
        for sheet in xl_workbook.sheet_names():
            xl_sheet = xl_workbook.sheet_by_name(sheet)
            sheet_data={}
            if sheet in sheetnames:
                for row_idx in range(1, xl_sheet.nrows):
                    row_val=[]
                    for col_idx in range(0,xl_sheet.ncols):
                        row_val.append((xl_sheet.cell(0,col_idx).value,xl_sheet.cell(row_idx,col_idx).value))
                    sheet_data[row_idx-1]=row_val
            data[sheet]=sheet_data
        return data

            
class DataFileWriter (object):
    
    def __init__(self,filepath,filename):
        self.filepath=filepath
        self.filename=filename
        self.file=join(self.filepath,self.filename)
    

    def delimitedfilewriter (self,data,delimiter=','):
        print (type(data[0]))
        with open(self.file, 'w',newline='') as out_file:
            writer = csv.writer(out_file, delimiter=delimiter)
            writer.writerow ( [val for val in data[0].keys()])
            for row,value in data.items():    
                writer.writerow (value.values())               
                
    def excelfilewriter (self,data,sheetnames=None):
        import xlwt
        workbook = xlwt.Workbook()
        for key,item in data.items():
            if sheetnames is None or key in sheetnames:
                sheet=workbook.add_sheet(key)
                fieldnames = [value[0] for value in item[0]]
                col_num=0
                for field in fieldnames:
                    sheet.write(0,col_num,field)
                    col_num += 1
                row_num=1           
                for key,value in item.items():
                    col_num=0
                    for val in value:
                        sheet.write(row_num,col_num,str((val[1])))
                        col_num += 1
                    row_num += 1
        workbook.save(self.file)
                

class TextFileReader (object):

        def __init__(self,filepath,filename):
            self.filepath=filepath
            self.filename=filename
            self.file=join(self.filepath,self.filename)
        

        
        def delimitedfilereader (self,delimiter,chunk_size=1024):
            
            with open (self.file,'r') as f:
                obj=FileSplitter(f)
                return [piece.strip().split(delimiter) for piece in obj.read_in_chunks()]

        def idelimitedfilereader (self,delimiter,chunk_size=1024):
            with open (self.file,'r') as f:
                obj=FileSplitter(f)
                for piece in obj.read_in_chunks():
                    yield piece.strip().split(delimiter)

                
            
            

if __name__ == '__main__':
    pass
    
#    obj=TextFileReader('E:\python\data','sample.csv')
#    print ([row for row in obj.idelimitedfilereader(',')])
#    exl_obj=DataFileReader('E:\python\data','sample.xlsx')
#    print (exl_obj.excelfilereader())
#    csv_obj=DataFileReader('E:\python\data','sample.csv')
#    print (csv_obj.delimitedfilereader(','))
#    write_csv_obj=DataFileWriter('E:\python','sample.csv')
#    write_csv_obj.delimitedfilewriter(csv_obj.delimitedfilereader(','),'|')
#    write_xls_obj=DataFileWriter('E:\python','sample.xls')
#    write_xls_obj.excelfilewriter(exl_obj.excelfilereader(),['Sheet2'])
    