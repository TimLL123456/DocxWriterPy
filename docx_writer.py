import zipfile
import os
from lxml import etree
import shutil
import win32com.client

class DocxWriter:
    def __init__(self, docx_path, temp_dir = "temp_docx"):
        """
        Initialize the DocxWriter with the path to the DOCX file and a temporary directory.
        
        Parameters:
        - docx_path: Path to the DOCX file to be processed.
        - temp_dir: Temporary directory for unzipping and modifying DOCX content (default is "temp_docx").
        """

        self.docx_path = docx_path
        assert os.path.isabs(self.docx_path), f"Please input absolute path of {self.docx_path}"

        self.temp_dir = temp_dir
        self.document_xml_path = os.path.join(self.temp_dir, "word", "document.xml")
        self.media_path = os.path.join(self.temp_dir, "word", "media")

        ### Unzip .docx file
        self._unzip_docx()

        ### Load docx content
        self._load()

    def _unzip_docx(self):
        """Unzip the DOCX file into a temporary directory."""
        
        ### Create a temporary directory to extract files if it does not exist
        if not os.path.exists(self.temp_dir):
            os.makedirs(self.temp_dir, exist_ok=True)

        ### Unzip the docx file using the zipfile module
        with zipfile.ZipFile(self.docx_path, "r") as zip_ref:
            zip_ref.extractall(self.temp_dir)

    def _load(self):
        """Load the document.xml and parse it."""
        
        ### Parse the document.xml file using lxml's etree
        self.tree = etree.parse(self.document_xml_path)
        self.root = self.tree.getroot()

        ### Extract namespaces for XML parsing
        self.namespaces = {"w": self.root.nsmap["w"]}

    @property
    def xml_tree(self):
        """Return the parsed XML tree."""
        return self.tree

    @property
    def xml_root(self):
        """Return the root element of the XML tree."""
        return self.root

    @property
    def xml_namespaces(self):
        """Return the namespaces used in the XML document."""
        return self.namespaces

    @property
    def paragraph(self):
        """Return a list of paragraphs as a single string joined by newlines."""
        
        para = []
        
        ### Find all paragraph elements in the XML document
        p_lst = self.root.findall(".//w:p", self.namespaces)
        
        for p in p_lst:
            curr_para = []
            ### Extract text from each text element within the paragraph
            for t in p.findall(".//w:t", self.namespaces):
                curr_para.append(t.text)        
            
            ### Join text elements to form a complete paragraph and add to list
            para.append("".join(curr_para))
        
        return "\n".join(para)

    @property
    def texts(self):
        """Return all text elements in the document."""
        
        return self.root.findall(".//w:t", self.namespaces)
    
    @property
    def textbox(self):
        """Return all text content from textboxes in the document."""
        
        txtxContent = self.root.findall(".//w:txbxContent", self.namespaces)
        
        ### Extract text from each textbox content element
        return [t for tbc in txtxContent for t in tbc.findall(".//w:t", self.namespaces)]
    
    def extract_image(self, output_path):
        """Extract images from the DOCX file to the specified output path."""
        
        ### Create output directory if it does not exist
        if not os.path.exists(output_path):
            os.makedirs(output_path, exist_ok=True)
        
        ### Copy each image from media folder to output path
        for img_name in os.listdir(self.media_path):
            img_path = os.path.join(self.media_path, img_name)
            new_img_path = os.path.join(output_path, img_name)
            shutil.copy(img_path, new_img_path)

    def text_replace(self, texts_list, origi_text, new_text, symbol=None):
        """
        Replace specified text within a list of text elements.
        
        Parameters:
        - texts_list: List of text elements to search through.
        - origi_text: Original text to be replaced.
        - new_text: New text that will replace the original.
        - symbol: Optional symbol that indicates specific replacement conditions.
        """
        
        for _index in range(len(texts_list)):
            if (_index == 0 or _index > len(texts_list) - 1):
                continue

            ### Replace without symbols if no symbol is provided
            if texts_list[_index].text == origi_text and symbol is None:
                texts_list[_index].text = texts_list[_index].text.replace(origi_text, new_text)

            ### Replace with symbols if specified
            if texts_list[_index].text == f"{symbol}{origi_text}{symbol}":
                texts_list[_index].text = texts_list[_index].text.replace(f"{symbol}{origi_text}{symbol}", new_text)

            if texts_list[_index].text == origi_text and \
                symbol is not None and \
                texts_list[_index-1].text == symbol and texts_list[_index+1].text == symbol:
                
                texts_list[_index].text = texts_list[_index].text.replace(origi_text, new_text)

                ### Remove surrounding symbols after replacement
                texts_list[_index-1].getparent().remove(texts_list[_index-1])
                texts_list[_index+1].getparent().remove(texts_list[_index+1])

    def image_replace(self, new_img_path, origi_img):
        """Replace an image in the DOCX file with a new image."""
        
        shutil.copy(new_img_path, os.path.join(self.media_path, origi_img))

    def save(self, new_docx_path):
        """Save changes made to the document and repack the temporary directory into a new DOCX file."""
        
        ### Write changes back to document.xml 
        self.tree.write(self.document_xml_path)

        ### Create a new DOCX file by zipping the contents of the temporary directory
        with zipfile.ZipFile(new_docx_path, "w") as new_docx:
            for foldername, subfolders, filenames in os.walk(self.temp_dir):
                for filename in filenames:
                    file_path = os.path.join(foldername, filename)
                    rel_path = os.path.relpath(file_path, self.temp_dir)
                    new_docx.write(file_path, rel_path)

    def close(self):
        """Remove temporary folder used during processing."""
        
        ### Walk through temp directory and remove files and folders
        for foldername, subfolders, filenames in os.walk(self.temp_dir, topdown=False):
            for filename in filenames:
                os.remove(os.path.join(foldername, filename))
            os.rmdir(foldername)

    def save_as_pdf(self, old_pdf_path, new_pdf_path):
        
        assert os.path.isabs(new_pdf_path), "Please input absolute path of the output pdf path"

        word = win32com.client.Dispatch("Word.Application")
        doc = word.Documents.Open(old_pdf_path)
        doc.SaveAs(new_pdf_path, FileFormat=17)
        doc.Close()
        # word.Quit() ### <-- This line of code may cause `Error` - `com_error: (-2147417848, '用戶端中斷了已啟動物件的連線。', None, None)`
