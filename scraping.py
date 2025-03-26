import time
import numpy as np
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains

driver = webdriver.Chrome()
driver.maximize_window()

wait= WebDriverWait(driver,5)

def wait_for_page_toload(driver, wait):
    page_title = driver.title
    try:
        wait.until(
            lambda d: d.excute_script("return ducument.readyState") == "complete"
        )
    except:
        print(f"{page_title} page not fully loaded.\n")
    else:
        print(f"{page_title} page has successfully loaded.\n")

url = "https://finance.yahoo.com/"
driver.get(url)
wait_for_page_toload(driver, wait)

## Accessing Market Menu

actions = ActionChains(driver)
market_menu = wait.until(
    EC.presence_of_element_located((By.XPATH,'/html[1]/body[1]/div[2]/header[1]/div[1]/div[1]/div[1]/div[4]/div[1]/div[1]/ul[1]/li[3]/a[1]/span[1]'))
)
actions.move_to_element(market_menu).perform()

## Clicking Trending_tickers Menu

trending_tickers = wait.until(
    EC.element_to_be_clickable((By.XPATH,'/html[1]/body[1]/div[2]/header[1]/div[1]/div[1]/div[1]/div[4]/div[1]/div[1]/ul[1]/li[3]/div[1]/ul[1]/li[4]/a[1]/div[1]'))
)
trending_tickers.click()
wait_for_page_toload(driver, wait)

## Clicking 52 week gainers Menu

weeks52_gainers = wait.until(
    EC.element_to_be_clickable((By.XPATH,'/html[1]/body[1]/div[2]/main[1]/section[1]/section[1]/section[1]/article[1]/section[1]/div[1]/nav[1]/ul[1]/li[5]/a[1]/span[1]'))
)
weeks52_gainers.click()
wait_for_page_toload(driver, wait)


data = []
while True:
    wait.until(
        EC.presence_of_element_located((By.TAG_NAME,"table"))
    )
    rows = driver.find_elements(By.CSS_SELECTOR, "table tbody tr")
    for row in rows:
        values = row.find_elements(By.TAG_NAME,"td")
        stocks = {
            "name" : values[1].text,
            "symbol" : values[0].text,
            "price" : values[3].text,
            "change" : values[4].text,
            "volume" : values[6].text,
            "market_cap": values[8].text,
            "pe_ratio": values[9].text,
        }
        data.append(stocks)
    try: 
        nxt_btn = wait.until(
        EC.element_to_be_clickable((By.XPATH,"//button[@aria-label='Goto next page']//div[@class='icon fin-icon inherit-icn sz-medium yf-z7spfk']//*[name()='svg']"))
        )
        
    except:
        print("All page navigated.")
        break
    else :
        nxt_btn.click()

#Clean Data
stocks_df = (
    pd.DataFrame(data)
        .assign(
            price=lambda df_: pd.to_numeric(df_.price),
            change=lambda df_: pd.to_numeric(df_.change.str.replace("+", "")),
         
            volume=lambda df_: df_.volume.apply(
                lambda val: float(val.replace(",", "").replace("M", "")) if isinstance(val, str) and "M" in val
                else float(val.replace(",", "")) / 1_000_000 if isinstance(val, str)
                else val  # Keeps NaN or numeric values unchanged
            ),

            market_cap=lambda df_: df_.market_cap.apply(lambda val: float(val.replace("M", "")) if "M" in val else float(val.replace("B", "")) * 1000),
            
            pe_ratio=lambda df_: (
			df_
			.pe_ratio
			.replace("-", np.nan)
			.str.replace(",", "")
			.pipe(lambda col: pd.to_numeric(col))
    		)
            
        )
            
        .rename(columns={
            "price":"price_usd",
            "market_cap":"market_cap_B",
            "volume":"volume_in_million"
            
        })
)
stocks_df
stocks_df.to_excel("yahoo12.xlsx",index=False)


driver.quit()
