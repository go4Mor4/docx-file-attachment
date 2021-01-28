from src.docx_writer import Docx

if __name__ == '__main__':
    docx = Docx("input/caat_compact.docx")
    response = docx._fill_logs()

    print(response)
