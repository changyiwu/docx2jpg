import os
import sys
import win32com.client
import fitz  # PyMuPDF

def convert_docx_to_jpg(docx_path):
    docx_path = os.path.abspath(docx_path)
    if not os.path.exists(docx_path):
        print(f"Error: The file {docx_path} does not exist.")
        sys.exit(1)
        
    output_dir = os.path.dirname(docx_path)
    base_name = os.path.splitext(os.path.basename(docx_path))[0]
    pdf_path = os.path.join(output_dir, f"{base_name}_temp.pdf")
    
    print(f"Opening Word document: {docx_path}")
    
    # Initialize Word
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    
    try:
        doc = word.Documents.Open(docx_path)
        # Save as PDF (FileFormat=17)
        print("Converting to temporary PDF...")
        doc.SaveAs(pdf_path, FileFormat=17)
        doc.Close()
        print("Word conversion finished.")
    except Exception as e:
        print(f"Error converting Word file: {e}")
        word.Quit()
        sys.exit(1)
    finally:
        word.Quit()
        
    # Convert PDF to JPG
    print(f"Converting PDF to JPG...")
    try:
        pdf_doc = fitz.open(pdf_path)
        num_pages = len(pdf_doc)
        print(f"Total pages: {num_pages}")
        
        generated_files = []
        for i, page in enumerate(pdf_doc):
            # Increase zoom for higher quality (3.0x zoom gives ~216 DPI)
            zoom = 3.0
            mat = fitz.Matrix(zoom, zoom)
            pix = page.get_pixmap(matrix=mat)
            
            if num_pages == 1:
                jpg_path = os.path.join(output_dir, f"{base_name}.jpg")
            else:
                jpg_path = os.path.join(output_dir, f"{base_name}_page{i+1}.jpg")
                
            pix.save(jpg_path)
            generated_files.append(jpg_path)
            print(f"Saved: {jpg_path}")
            
        pdf_doc.close()
    except Exception as e:
        print(f"Error converting PDF to JPG: {e}")
        sys.exit(1)
    finally:
        # Clean up temporary PDF
        if os.path.exists(pdf_path):
            os.remove(pdf_path)
            print("Temporary PDF removed.")
            
    print("\nConversion successfully completed!")
    print("Generated files:")
    for f in generated_files:
        print(f" - {f}")

if __name__ == "__main__":
    target_file = r"H:\我的雲端硬碟\進修部\教務\段考日程\114\第2學期第3次段考日程表.docx"
    convert_docx_to_jpg(target_file)
