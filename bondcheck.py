import time
import os
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select, WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# Load Excel file with bond numbers
file_path = "Testing_data.xlsx"
df = pd.read_excel(file_path)

# Initialize Selenium WebDriver
options = webdriver.ChromeOptions()
options.add_argument("--start-maximized")
driver = webdriver.Chrome(options=options)
driver.get("https://hamariweb.com/finance/prizebonds/")

# Prepare the output Excel file
output_file = "results.xlsx"
if os.path.exists(output_file):
    os.remove(output_file)

# Wait object
wait = WebDriverWait(driver, 15)

# Loop through each column (bond amount)
for amount in df.columns:
    amount_value = str(int(float(amount)))  # e.g., "100.00" -> "100"

    try:
        # Select bond amount from dropdown
        dropdown = wait.until(
            EC.presence_of_element_located((By.ID, "PageContent_ddPB"))
        )
        Select(dropdown).select_by_value(amount_value)
        time.sleep(2)

        # Loop through all bond numbers in this column
        for bond_number in df[amount].dropna():
            try:
                bond_str = str(int(bond_number)).zfill(6)  # Ensure 6-digit format

                input_field = wait.until(
                    EC.presence_of_element_located((By.ID, "txtNumber"))
                )
                input_field.clear()
                input_field.send_keys(bond_str)
                time.sleep(1)

                add_btn = wait.until(
                    EC.element_to_be_clickable((By.CLASS_NAME, "btn_add"))
                )
                add_btn.click()
                time.sleep(2)

                print(f"Added Bond: {bond_str} for Amount: {amount_value}")
            except Exception as e:
                print(f"Error adding bond {bond_number} for Amount {amount_value}: {e}")

        # Click the 'Check' button
        try:
            check_btn = wait.until(
                EC.element_to_be_clickable((By.CLASS_NAME, "btn_check"))
            )
            check_btn.click()
            time.sleep(5)
        except Exception as e:
            print(f"Error clicking Check for Amount {amount_value}: {e}")

        # Extract results
        data = []
        try:
            result_div = wait.until(EC.presence_of_element_located((By.ID, "result")))
            time.sleep(2)

            try:
                table = result_div.find_element(By.TAG_NAME, "table")
                rows = table.find_elements(By.TAG_NAME, "tr")
                for row in rows:
                    cells = row.find_elements(By.TAG_NAME, "td")
                    if cells:
                        row_data = [cell.text.strip() for cell in cells]
                        data.append(row_data)
            except Exception:
                try:
                    h4_elem = result_div.find_element(By.TAG_NAME, "h4")
                    data.append([h4_elem.text.strip()])
                except Exception:
                    result_text = result_div.text.strip()
                    if result_text:
                        data.append([result_text])
                    else:
                        data.append(["No result found."])
        except Exception as e:
            print(f"Could not extract results for Amount {amount_value}: {e}")

        # Save results
        try:
            df_result = pd.DataFrame(data)
            with pd.ExcelWriter(
                output_file,
                engine="openpyxl",
                mode="a" if os.path.exists(output_file) else "w",
            ) as writer:
                df_result.to_excel(
                    writer, sheet_name=f"Amount_{amount_value}", index=False
                )
            print(f"Results saved for Amount: {amount_value}")
        except Exception as e:
            print(f"Could not save results for Amount {amount_value}: {e}")

        # Clear results for next iteration
        try:
            clear_btn = wait.until(
                EC.element_to_be_clickable((By.CLASS_NAME, "btn_clear"))
            )
            clear_btn.click()
            time.sleep(2)
        except Exception:
            print("Clear button not found, clearing result div via JS...")
            driver.execute_script("document.getElementById('result').innerHTML = '';")
            time.sleep(2)

    except Exception as e:
        print(f"Unexpected error processing Amount {amount_value}: {e}")

# Final message
print("All columns processed.")
driver.quit()
print("Process completed.")
