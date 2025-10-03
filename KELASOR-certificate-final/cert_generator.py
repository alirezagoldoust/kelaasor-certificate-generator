from bs4 import BeautifulSoup
import cairosvg
import os
from colorama import Fore, Style
import requests
from dotenv import load_dotenv

BASE_DIRECTORY = os.path.dirname(os.path.abspath(__file__))
print(BASE_DIRECTORY)

load_dotenv(f'{BASE_DIRECTORY}/.env')


def get_google_sheet_students_data():
    try:
        API_KEY = os.getenv("GOOGLE_SHEET_API_KEY")
        SHEET_ID = os.getenv("GOOGLE_SHEET_ID")
        RANGE = "Students" 
        url = f"https://sheets.googleapis.com/v4/spreadsheets/{SHEET_ID}/values/{RANGE}?key={API_KEY}&majorDimension=ROWS"

        # Fetch data
        response = requests.get(url)
        response.raise_for_status()
        data_list = response.json().get("values", [])
        print(f"{Fore.GREEN}✔{Style.RESET_ALL} Data fetched successfully from Google Sheets.\n")
        return data_list

    except Exception as e:
        print(f"{Fore.RED}‼️{Style.RESET_ALL} Error fetching data from Google Sheets: {e}")
        return []

def generate_certificate(student_first_name, student_last_name, student_pre_name, student_bootcamp):
    html_path = f'{BASE_DIRECTORY}/cert.html'
    output_html_path = f'{BASE_DIRECTORY}/cert_tmp.html'
    student_full_name = f"{student_first_name} {student_last_name}"
    output_format='png'
    with open(html_path, 'r', encoding='utf-8') as file:
        html_content = file.read()
        
    possessive_adjective = "his" if student_pre_name.lower() == "mr" else "her"

    # Parse SVG with BeautifulSoup
    soup = BeautifulSoup(html_content, 'xml')

    # Find and replace the text content
    # Header
    title_element = soup.find_all(id='student-header-name')[0]
    title_element.string = student_full_name
    
    # inner texts
    # name_element = soup.find_all(id='student-text-name')[0]
    # name_element.string = f"{pre_name} {student_name}"
    
    name_element = soup.find_all(id='student-text-name')[0]
    name_element.string = f"{student_pre_name}. {student_full_name}"
        
    bootcamp_element = soup.find_all(id='bootcamp-name')[0]
    bootcamp_element.string = student_bootcamp
    
    formal_name_element = soup.find_all(id='student-text-formal-name')[0]
    formal_name_element.string = f"{student_pre_name}.{student_last_name}"
    
    adjectives = soup.find_all(class_='student-adjective')
    for adjective in adjectives: 
        adjective.string = possessive_adjective
    
    with open(f"{output_html_path}", 'w', encoding='utf-8') as file:
        file.write(str(soup.contents[0]))

    # Create output directory if it doesn't exist
    output_dir = f"{'/'.join(BASE_DIRECTORY.split('/')[:-1])}/generated_certificates"
    os.makedirs(output_dir, exist_ok=True)

    # Generate output filename
    output_filename = f"{student_full_name.replace(' ', '_').lower()}_certificate.{output_format}"
    output_path = os.path.join(output_dir, output_filename)

    # Convert modified HTML to PNG
    from playwright.sync_api import sync_playwright

    with sync_playwright() as p:
        browser = p.chromium.launch()
        page = browser.new_page(viewport={"width": 1920, "height": 1080}) 
        page.goto(f"file://{output_html_path}")
        page.screenshot(path=output_path, full_page=True, scale="device")
        browser.close()

    return output_path


def generate_from_terminal():
    print(f"\n\n{Style.BRIGHT}{Fore.BLUE}======= Kelaasor Certificate Generator ======={Style.RESET_ALL}\n")
    student_first_name = input(f"- Please enter student's {Style.BRIGHT}{Fore.GREEN}first name: {Style.RESET_ALL}")
    student_last_name = input(f"- Please enter student's {Style.BRIGHT}{Fore.GREEN}last name: {Style.RESET_ALL}")
    student_pre_name = input(f"- Please enter student's {Style.BRIGHT}{Fore.GREEN}pre name(Mr or Ms or Mrs): {Style.RESET_ALL}")
    student_bootcamp = input(f"- Please enter student's {Style.BRIGHT}{Fore.GREEN}bootcamp: {Style.RESET_ALL}")


    output_path = generate_certificate(student_first_name, student_last_name, student_pre_name, student_bootcamp)
    print(f"\n✅ Certificate generated in: {output_path}")

def generate_from_google_sheet():
    """
    standard_data would be like:
    [
        ['student_first_name',   'student_last_name', 'student_pre_name', 'bootcamp_name'],
        ['alireza',              'goldoust',          'Mr',               'Back-end'],
    ]
    
    """
    print(f"\n\n{Style.BRIGHT}{Fore.BLUE}======= Kelaasor Certificate Generator ======={Style.RESET_ALL}\n")
    print("        getting data from google sheet...\n\n")
    students = get_google_sheet_students_data()[1:]
    for student in students:
        generate_certificate(*student)
        print(f"✅ Certificate generated for {student[0]} {student[1]}")

generate_from_google_sheet()