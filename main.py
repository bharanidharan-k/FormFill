import os, platform
import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import Select

df = pd.read_excel("./nism_registration.xlsx")
df = df[df.to_be_registered=='Yes']
registration_data = df.to_dict(orient='records')


def uploadImage(driver, data, imgName, id):
    try:
        imagePath = os.getcwd()
        if platform.system() == 'Windows':
            imagePath += '\\' + data['firstName'].strip() + '-' + data['pan'] + '\\' + data['firstName'].strip() + imgName
        else:
            imagePath += '/' + data['firstName'].strip() + '-' + data['pan'] + '/' + data['firstName'].strip() + imgName
        if os.path.exists(imagePath+'.jpeg'):
            imagePath += '.jpeg'
        else:
            imagePath += '.jpg'
        print(imagePath)

        if os.path.exists(imagePath):
            driver.find_element_by_id(id).click()
            driver.switch_to.frame(0);
            driver.find_element_by_id('flpPhoto').send_keys(imagePath)
            time.sleep(2)
            driver.find_element_by_id('btnUpload').click()
            s = driver.find_element_by_class_name('jcrop-holder').get_attribute('style')
            w = int(s[s.index("width: ")+7:s.index("px;")])
            h = int(s[s.index("height: ")+8:s.rindex("px;")])
            action = ActionChains(driver)
            action.click_and_hold(on_element = driver.find_element_by_class_name('jcrop-tracker'))
            action.move_by_offset(w, h)
            action.release()
            action.perform()
            driver.find_element_by_id('btnCrop').click()
            time.sleep(4)
            driver.find_element_by_id('btnSubmit').click()
            time.sleep(5)
    except:
        pass


def fillGender(driver, data):
    try:
        if data['gender'] == 'M' or data['gender'].lower() == 'male':
            Select(driver.find_element_by_id('ddlTitle')).select_by_index(1)
            Select(driver.find_element_by_id('ddlGender')).select_by_index(1)

        else:
            Select(driver.find_element_by_id('ddlTitle')).select_by_index(2)
            Select(driver.find_element_by_id('ddlGender')).select_by_index(2)
    except:
        pass


def fillContactDetails(driver, data):
    try:
        address = data['address']
        address_lines = address.split(',')
        Select(driver.find_element_by_id('ddlCountry')).select_by_index(1)
        driver.find_element_by_id('txtAddress1').send_keys((address_lines[0] + address_lines[1]).strip())
        driver.find_element_by_id('txtAddress2').send_keys(address_lines[2].strip())
        driver.find_element_by_id('txtAddress3').send_keys(address_lines[-1].strip())
        driver.find_element_by_id('txtCity').send_keys(data['city'])
        driver.find_element_by_id('txtPinCode').send_keys(data['pincode'])
        Select(driver.find_element_by_id('ddlState')).select_by_index(32)
        driver.find_element_by_id('txtMobile').send_keys(data['mobile'])
    except:
        pass


def fillPersonalDetails(driver, data):
    try:
        driver.find_element_by_id('txtEmailId').send_keys(data['email'])
        driver.find_element_by_id('txtConfirmEmailId').send_keys(data['email'])
        driver.find_element_by_id('txtPassword').send_keys(data['password'])
        driver.find_element_by_id('txtConfirmPassword').send_keys(data['password'])
        fillGender(driver, data)
        driver.find_element_by_id('txtFirstName').send_keys(data['firstName'])
        driver.find_element_by_id('txtFatherName').send_keys(data['fatherName'])
        driver.find_element_by_id('txtDateOfBirth').send_keys(data['dob'])
        driver.find_element_by_id('txtPan').send_keys(data['pan'])
        driver.find_element_by_id('txtConfirmPan').send_keys(data['pan'])
        uploadImage(driver, data, 'PP', 'imgUserPhoto')
        uploadImage(driver, data, 'PAN', 'imgPAN')
        uploadImage(driver, data, 'AADHAAR', 'imgPOA')
        driver.find_element_by_id('strAadhaarNumber').send_keys(data['aadhaar'])
        driver.find_element_by_id('txtLastName').send_keys(data['lastName'])
        driver.execute_script('document.getElementById("txtDateOfBirth").removeAttribute("readonly")')
        driver.find_element_by_id('txtDateOfBirth').send_keys(data['dob'])
    except:
        pass


def fillEducationDetails(driver, data):
    try:
        if data['education'].lower() == 'graduation':
            Select(driver.find_element_by_id('ddlQualification')).select_by_index(3)
            if 'commerce' in data['major'].lower():
                Select(driver.find_element_by_id('ddlMajorSubject')).select_by_index(3)
            elif 'b.sc' in data['major'].lower() or 'bsc' in data['major'].lower():
                Select(driver.find_element_by_id('ddlMajorSubject')).select_by_index(2)
            elif 'engineer' in data['major'].lower() or 'engg' in data['major'].lower() or 'b.tech' in data['major'].lower() or 'b.e' in data['major'].lower():
                Select(driver.find_element_by_id('ddlMajorSubject')).select_by_index(4)
            elif 'bba' in data['major'].lower() or 'mba' in data['major'].lower():
                Select(driver.find_element_by_id('ddlMajorSubject')).select_by_index(7)
            driver.find_element_by_id('txtGradInstitute').send_keys(data['university'])
            Select(driver.find_element_by_id('ddlGradYearofPassing')).select_by_visible_text(str(data['yop']))
            driver.find_element_by_id('txtGradPercenteage').send_keys(str(data['percentage']))
    except:
        pass


def fillExperienceDetails(driver, data):
    try:
        if data['occupation'].lower() == 'employed':
            Select(driver.find_element_by_id('ddlOccupation')).select_by_index(2)
            if data['major'].lower() == 'commerce':
                Select(driver.find_element_by_id('ddlMajorSubject')).select_by_index(3)
            if data['country_office'].lower() == 'india':
                Select(driver.find_element_by_id('ddlOfficeCountry')).select_by_visible_text(data['country_office'])
            driver.find_element_by_id('txtNameOfOrganization').send_keys(data['organization'])
            driver.find_element_by_id('txtDesignation').send_keys(data['designation'])
            # driver.find_element_by_id('txtFromDate').send_keys(data['from_date'].strftime("%d/%m/%Y"))
            address_lines = data['address_office'].split(',')
            driver.find_element_by_id('txtOfficeAdd1').send_keys((address_lines[0]+address_lines[1]).strip())
            driver.find_element_by_id('txtOfficeAdd2').send_keys(address_lines[2].strip())
            driver.find_element_by_id('txtOfficeAdd3').send_keys(address_lines[-1].strip())
            driver.find_element_by_id('txtOfficeCity').send_keys(data['city_office'])
            driver.find_element_by_id('txtOfficePinCode').send_keys(data['pincode_office'])
            if data['phone']:
                driver.find_element_by_id('txtEmpSTD').send_keys(data['phone'].split('-')[0])
                driver.find_element_by_id('txtEmpPhoneNo').send_keys(data['phone'].split('-')[1])
            if data['country_office'].lower() == 'india':
                Select(driver.find_element_by_id('ddlOfficeState')).select_by_visible_text(data['state_office'])
    except:
        pass


driver = webdriver.Chrome('./chromedriver')

for i in range(len(registration_data)):
    driver.get("https://certifications.nism.ac.in/nismaol/Candidate/UserRegistration.aspx")
    data = registration_data[i]
    fillPersonalDetails(driver, data)
    fillContactDetails(driver, data)
    fillEducationDetails(driver, data)
    fillExperienceDetails(driver, data)
    if i < len(registration_data)-1:
        driver.execute_script(f"window.open('about:blank', 'tab'+{i});")
        driver.switch_to.window(f"tab{i}")
