from Scrap import (check_robots_txt_and_sitemap,  get_speed_index, domain_rating, convert_to_grade, count_numbers_less_than_10, scroll_to_bottom, scroll_to_element, scrape_data, get_chatgpt_response, add_heading, add_bullet, add_bullet_bold, generate_seo_report)  # Import the functions from seo_report.py
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
import time
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import TimeoutException
import undetected_chromedriver as uc
import re
from docx import Document
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.shared import Pt, Inches, RGBColor
from docx.oxml import OxmlElement  # Ensure this import is included
import pandas as pd
import openai
from fastapi import FastAPI, HTTPException, BackgroundTasks
from pydantic import BaseModel
from fastapi.responses import FileResponse
import requests
import os

app = FastAPI()

REPORTS_DIR = "static/reports"
os.makedirs(REPORTS_DIR, exist_ok=True)

class SEOReportRequest(BaseModel):
    company_name: str
    company_url: str
    company_xpath: str

def generate_report_task(company_name, company_url, company_xpath):
    """Runs scraping and SEO report generation in a background task."""
    try:
        # Step 1: Scrape data and get CSV file
        csv_file_path = scrape_data(company_name, company_url, company_xpath)

        # Step 2: Generate SEO report (Word document) using the scraped CSV data
        report_name = generate_seo_report(csv_file_path, company_name)

        report_path = os.path.join(REPORTS_DIR, report_name)
        os.rename(report_name, report_path)

        print(f"Report generated: {report_path}")
        return report_name # Return path of the generated report
    except Exception as e:
        raise Exception(f"Error: {str(e)}")

@app.post("/scraped/")
async def generate_report(request: SEOReportRequest, background_tasks: BackgroundTasks):
    try:
        # Run the report generation as a background task
        background_tasks.add_task(generate_report_task, request.company_name, request.company_url, request.company_xpath)

        expected_filename = f"{request.company_name}.docx"
        download_url = f"/download/{expected_filename}"

        return {
            "message": "Report generation started. Download link will be available shortly.",
            "download_link": download_url
        }

    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Error: {str(e)}")

@app.get("/download/{filename}")
async def download_report(filename: str):
    """Endpoint to download the generated Word document."""
    file_path = os.path.join(REPORTS_DIR, filename)

    if not os.path.exists(file_path):
        raise HTTPException(status_code=404, detail="File not found.")

    return FileResponse(file_path, media_type='application/vnd.openxmlformats-officedocument.wordprocessingml.document', filename=filename)
