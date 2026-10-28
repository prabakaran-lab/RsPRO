import urllib.request
import urllib.error
import json
import time
import os

url_upload = "http://127.0.0.1:8001/api/job/upload"
excel_file = r"c:\Users\Kalai\Downloads\XSLX TO PDF\XSLX\Speciality Size Rechargeable Batteries (3) (1).xlsx"

if not os.path.exists(excel_file):
    print(f"Error: Excel file not found at {excel_file}")
    exit(1)

print(f"Testing Speciality XLSX Upload with file: {excel_file}")

boundary = '----WebKitFormBoundary7MA4YWxkTrZu0gW'
data = []
data.append(f'--{boundary}'.encode('utf-8'))
data.append(f'Content-Disposition: form-data; name="file"; filename="{os.path.basename(excel_file)}"'.encode('utf-8'))
data.append(f'Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'.encode('utf-8'))
data.append(b'')
with open(excel_file, 'rb') as f:
    data.append(f.read())
data.append(f'--{boundary}--'.encode('utf-8'))
data.append(b'')

body = b'\r\n'.join(data)

req = urllib.request.Request(url_upload, data=body)
req.add_header('Content-Type', f'multipart/form-data; boundary={boundary}')
req.add_header('Content-Length', str(len(body)))

try:
    print("Uploading Excel sheet to server...")
    with urllib.request.urlopen(req) as response:
        res_body = response.read().decode('utf-8')
        job = json.loads(res_body)
        job_id = job['id']
        print(f"Job successfully created! Job ID: {job_id} | Total Rows: {job['totalRows']}")
        
    # Poll job status
    print("Polling job progress...")
    for _ in range(30):
        time.sleep(3)
        status_req = urllib.request.Request(f"http://127.0.0.1:8001/api/job/{job_id}")
        with urllib.request.urlopen(status_req) as status_resp:
            job_status = json.loads(status_resp.read().decode('utf-8'))
            print(f"  Status: {job_status['status']} | Progress: {job_status['progress']}% | Success: {job_status['successCount']} | Failed: {job_status['failedCount']}")
            if job_status['status'] in ('COMPLETED', 'FAILED'):
                break
                
    # List generated PDFs
    print("\nListing generated PDFs in static/pdfs folder:")
    static_pdf_dir = r"c:\Users\Kalai\Downloads\XSLX TO PDF\static\pdfs"
    if os.path.exists(static_pdf_dir):
        files = os.listdir(static_pdf_dir)
        print(f"  Total generated files in folder: {len(files)}")
        for f in files:
            print(f"    - {f} ({os.path.getsize(os.path.join(static_pdf_dir, f))} bytes)")
    else:
        print("  Static PDF folder does not exist yet.")
        
except urllib.error.HTTPError as e:
    print(f"HTTP Error: {e.code}")
    print(e.read().decode('utf-8', errors='ignore'))
except Exception as e:
    print(f"Error: {e}")
