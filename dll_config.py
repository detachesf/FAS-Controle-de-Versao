import re
#Arquivo para tratamento das dlls capturadas através do listdll, que estão no arquivo Dlls_capturadas

class Config(object):
    """docstring for Config"""

    def __init__(self, input_file, content="msys64"):
        self.input_file_open = open(input_file, "r")
        self.content = [content, ".dll"]

    def separate_lines(self):
        output_file = open("Dlls.txt", "w")
        for row in self.input_file_open:
            if not False in [item in row for item in self.content]:
                output_file.write(row)
        output_file.close()

    def delete_duplicate_rows(self):
        Dlls_read = open("Dlls.txt", "r")
        Dlls = ""

        lines_seen = set()
        for row in Dlls_read:
            if row not in lines_seen:
                Dlls = Dlls + row
                lines_seen.add(row)

        Dlls_write = open("Dlls.txt", "w")
        Dlls_write.write(Dlls)

        Dlls_write.close()
        Dlls_read.close()

    def select_only_dlls(self):
        Dlls_read = open("Dlls.txt", "r")
        temp = ""
        temp = temp + "include_files = [\n"

        for row in Dlls_read:
            match = re.search(r'lld.*?[^\\]+', row[::-1])
            if match:
                temp = temp + "('" + row[:-1] + "', '" + match.group()[::-1] + "'),\n"

        temp = temp + "]"

        Dlls_write = open("Dlls.py", "w")
        temp = temp.replace("\\","/")
        print(temp)
        Dlls_write.write(temp)

        Dlls_write.close()
        Dlls_read.close()

    def run(self):
        self.separate_lines()
        self.delete_duplicate_rows()
        self.select_only_dlls()
        self.input_file_open.close()


if __name__ == '__main__':
    config = Config("Dlls_capturadas")
    config.run()
