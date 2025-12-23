from flask import Flask, request, render_template, send_file, jsonify, abort
import os
import tempfile
import subprocess
import sys
from werkzeug.utils import secure_filename

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024  # 50MB limit
app.config['UPLOAD_FOLDER'] = '/tmp'

# PDF to PPTX function using pdf2pptx[citation:1]
def convert_pdf_to_pptx(pdf_path, pptx_path):
    try:
        # Use pdf2pptx library
        from pdf2pptx import convert
        convert(pdf_path, pptx_path)
        return True
    except Exception as e:
        print(f"PDF to PPTX error: {e}")
        # Fallback method using LibreOffice if available
        try:
            subprocess.run(['libreoffice', '--headless', '--convert-to', 'pptx', 
                          '--outdir', os.path.dirname(pptx_path), pdf_path], 
                         check=True)
            return True
        except:
            return False

# PPTX to PDF function[citation:5][citation:8]
def convert_pptx_to_pdf(pptx_path, pdf_path):
    try:
        # Method 1: Using comtypes (Windows/Microsoft Office)
        try:
            import comtypes.client
            powerpoint = comtypes.client.CreateObject("Powerpoint.Application")
            powerpoint.Visible = 1
            deck = powerpoint.Presentations.Open(pptx_path)
            deck.SaveAs(pdf_path, FileFormat=32)  # 32 = PDF format
            deck.Close()
            powerpoint.Quit()
            return True
        except:
            # Method 2: Using LibreOffice (Linux)[citation:8]
            subprocess.run(['libreoffice', '--headless', '--convert-to', 'pdf', 
                          pptx_path, '--outdir', os.path.dirname(pdf_path)], 
                         check=True)
            return True
    except Exception as e:
        print(f"PPTX to PDF error: {e}")
        return False

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/convert', methods=['POST'])
def convert_file():
    try:
        if 'file' not in request.files:
            return jsonify({'error': 'No file uploaded'}), 400
        
        file = request.files['file']
        conversion_type = request.form.get('type', 'pdf2pptx')
        
        if file.filename == '':
            return jsonify({'error': 'No file selected'}), 400
        
        # Create temporary files
        with tempfile.NamedTemporaryFile(delete=False, suffix=os.path.splitext(file.filename)[1]) as temp_input:
            file.save(temp_input.name)
            input_path = temp_input.name
        
        # Determine output file
        if conversion_type == 'pdf2pptx':
            output_suffix = '.pptx'
            output_filename = 'converted_presentation.pptx'
            mimetype = 'application/vnd.openxmlformats-officedocument.presentationml.presentation'
        else:  # pptx2pdf
            output_suffix = '.pdf'
            output_filename = 'converted_document.pdf'
            mimetype = 'application/pdf'
        
        with tempfile.NamedTemporaryFile(delete=False, suffix=output_suffix) as temp_output:
            output_path = temp_output.name
        
        # Perform conversion
        success = False
        if conversion_type == 'pdf2pptx':
            success = convert_pdf_to_pptx(input_path, output_path)
        else:
            success = convert_pptx_to_pdf(input_path, output_path)
        
        if not success:
            return jsonify({'error': 'Conversion failed. Please check file format.'}), 500
        
        # Return the converted file
        return send_file(
            output_path,
            as_attachment=True,
            download_name=output_filename,
            mimetype=mimetype
        )
        
    except Exception as e:
        return jsonify({'error': str(e)}), 500
    finally:
        # Cleanup
        for path in [input_path, output_path]:
            if 'path' in locals() and os.path.exists(path):
                try:
                    os.unlink(path)
                except:
                    pass

@app.route('/health')
def health():
    return jsonify({'status': 'healthy'})

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 10000))
    app.run(host='0.0.0.0', port=port, debug=False)
