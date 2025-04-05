import os
import pandas as pd
import time
import random
import shutil
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
import openpyxl
from openpyxl_image_loader import SheetImageLoader

# 读取Excel文件
def read_excel(file_path):
    try:
        # 使用pandas读取数据
        df = pd.read_excel(file_path)
        print(f"成功读取Excel文件，共有{len(df)}行数据")
        return df
    except Exception as e:
        print(f"读取Excel文件时出错: {e}")
        return None

# 准备产品数据（替换原来的check_images函数）
def prepare_product_data(df, image_folder="product_images"):
    # 创建图片文件夹（如果不存在）
    if not os.path.exists(image_folder):
        os.makedirs(image_folder)
        print(f"创建图片文件夹: {image_folder}")
    
    # 添加图片路径列和图片状态列
    df['图片路径'] = ""
    df['有图片'] = False
    
    print("准备产品数据...")
    
    # 为每个产品生成图片路径
    for index, row in df.iterrows():
        try:
            # 获取产品信息
            brand = str(row['品牌']) if pd.notna(row['品牌']) else ""
            model = str(row['型号']) if pd.notna(row['型号']) else ""
            name = str(row['品名']) if pd.notna(row['品名']) else ""
            
            # 创建标准化的文件名
            brand_clean = brand.replace('/', '_').strip()
            model_clean = model.replace('/', '_').strip()
            name_clean = name.replace('/', '_').strip()
            
            # 生成图片文件名
            image_filename = f"{brand_clean}-{model_clean}-{name_clean}.jpg"
            
            # 生成图片路径
            image_path = os.path.join(image_folder, image_filename)
            
            # 保存图片路径
            df.at[index, '图片路径'] = image_path
            
            # 检查图片是否存在
            if os.path.exists(image_path):
                df.at[index, '有图片'] = True
                print(f"产品 '{brand} {model} {name}' 已有图片，将跳过上传")
            else:
                print(f"产品 '{brand} {model} {name}' 没有图片，将进行上传")
            
        except Exception as e:
            print(f"处理产品数据时出错 (行 {index+2}): {e}")
    
    print(f"已准备 {len(df)} 个产品的数据")
    print(f"其中 {df['有图片'].sum()} 个产品有图片（将跳过），{len(df) - df['有图片'].sum()} 个产品没有图片（将上传）")
    
    return df

# 创建中英文品名映射表
def create_name_mapping(df, mapping_file="name_mapping_new.xlsx"):
    # 确保不使用可能导致权限问题的文件名
    try:
        # 创建映射DataFrame
        unique_names = df['品名'].dropna().unique()
        mapping_df = pd.DataFrame({
            '中文品名': unique_names,
            '英文品名': [''] * len(unique_names)  # 空白，等待用户填写
        })
        
        # 保存映射表
        try:
            mapping_df.to_excel(mapping_file, index=False)
            print(f"已创建中英文品名映射表: {mapping_file}，请在此文件中填写对应的英文品名")
        except PermissionError:
            # 如果遇到权限错误，使用随机文件名
            new_filename = f"name_mapping_{random.randint(1000, 9999)}.xlsx"
            print(f"保存原文件时遇到权限错误，尝试另存为: {new_filename}")
            mapping_df.to_excel(new_filename, index=False)
            mapping_file = new_filename
            print(f"已创建中英文品名映射表: {mapping_file}，请在此文件中填写对应的英文品名")
        
        return mapping_file
    except Exception as e:
        print(f"创建映射表时出错: {e}")
        # 尝试创建CSV格式的备用映射表
        try:
            backup_file = "name_mapping_backup.csv"
            unique_names = df['品名'].dropna().unique()
            mapping_df = pd.DataFrame({
                '中文品名': unique_names,
                '英文品名': [''] * len(unique_names)
            })
            mapping_df.to_csv(backup_file, index=False, encoding='utf-8-sig')
            print(f"已创建备用CSV格式映射表: {backup_file}")
            return backup_file
        except Exception as e2:
            print(f"创建备用映射表也失败: {e2}")
            return None

# 读取填写好的映射表
def read_mapping(mapping_file):
    try:
        if mapping_file.endswith('.xlsx'):
            mapping_df = pd.read_excel(mapping_file)
        elif mapping_file.endswith('.csv'):
            mapping_df = pd.read_csv(mapping_file, encoding='utf-8-sig')
        else:
            print(f"不支持的映射文件格式: {mapping_file}")
            return {}
            
        # 创建字典
        name_map = dict(zip(mapping_df['中文品名'], mapping_df['英文品名']))
        print(f"成功读取映射表，共有{len(name_map)}个映射")
        return name_map
    except Exception as e:
        print(f"读取映射表时出错: {e}")
        return {}

# 使用Selenium上传产品到WordPress
def upload_to_wordpress(df, name_map, wp_url, username, password):
    # 添加更多的Selenium配置选项
    options = webdriver.ChromeOptions()
    options.add_argument("--start-maximized")  # 最大化窗口
    options.add_argument("--disable-extensions")  # 禁用扩展
    options.add_argument("--disable-gpu")  # 禁用GPU加速
    options.add_argument("--no-sandbox")  # 禁用沙盒模式
    options.add_argument("--disable-dev-shm-usage")  # 禁用/dev/shm使用
    
    print("正在初始化Chrome浏览器...")
    
    try:
        driver = webdriver.Chrome(options=options)
        print("Chrome浏览器已成功启动")
        
        # 确保URL格式正确
        if not wp_url.startswith(('http://', 'https://')):
            # 对于本地地址，使用http://前缀
            if wp_url.startswith(('localhost', '127.0.0.1')):
                wp_url = 'http://' + wp_url
            else:
                wp_url = 'https://' + wp_url
                
        # 移除URL末尾可能的/wp-admin部分，因为后面会添加
        if wp_url.endswith('/wp-admin'):
            wp_url = wp_url[:-9]
        elif '/wp-admin/' in wp_url:
            wp_url = wp_url.split('/wp-admin/')[0]
            
        print(f"使用的WordPress网址: {wp_url}")
        
        # 登录WordPress
        print("正在登录WordPress...")
        login_url = f"{wp_url}/wp-login.php"
        print(f"访问登录页面: {login_url}")
        driver.get(login_url)
        
        # 等待登录页面加载
        try:
            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, "user_login"))
            )
            print("登录页面已加载")
        except TimeoutException:
            print(f"无法加载登录页面: {login_url}")
            print(f"当前页面标题: {driver.title}")
            print(f"当前URL: {driver.current_url}")
            raise Exception("登录页面加载失败")
        
        # 输入登录信息
        driver.find_element(By.ID, "user_login").send_keys(username)
        driver.find_element(By.ID, "user_pass").send_keys(password)
        driver.find_element(By.ID, "wp-submit").click()
        
        # 等待登录完成
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "wpadminbar"))
        )
        print("登录成功")
        
        # 先导航到产品页面，确保完全进入后台
        print("导航到WordPress产品管理页面...")
        driver.get(f"{wp_url}/wp-admin/edit.php?post_type=product")
        
        # 等待产品管理页面加载完成
        try:
            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "a.page-title-action"))
            )
            print("已成功进入产品管理页面")
            time.sleep(2)  # 额外等待确保页面完全加载
        except TimeoutException:
            print("无法找到添加新产品按钮，尝试刷新页面...")
            driver.refresh()
            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "a.page-title-action"))
            )
            time.sleep(2)
        
        # 上传产品
        upload_count = 0
        for index, row in df.iterrows():
            try:
                # 在尝试访问数据前先定义变量，避免异常时引用未定义变量
                brand = ""
                model = ""
                price = ""
                chinese_name = ""
                english_name = ""
                image_path = ""
                current_operation = "获取产品基本信息"
                
                # 获取产品信息，添加更多的错误检查
                brand = str(row['品牌']) if pd.notna(row['品牌']) else ""
                model = str(row['型号']) if pd.notna(row['型号']) else ""
                
                # 从C列读取价格信息
                try:
                    # 获取C列的值作为价格
                    price = str(row.iloc[2]) if pd.notna(row.iloc[2]) else ""
                    # 如果价格为空，尝试使用单价字段作为备选
                    if not price and '单价' in row and pd.notna(row['单价']):
                        price = str(row['单价'])
                    print(f"从C列读取的价格: {price}")
                except Exception as price_error:
                    print(f"警告: 行 {index+2} 读取C列价格时出错: {price_error}")
                    # 尝试使用单价字段作为备选
                    if '单价' in row and pd.notna(row['单价']):
                        price = str(row['单价'])
                        print(f"使用单价字段作为备选: {price}")
                    else:
                        price = ""
                        print("无法获取价格信息，将使用空值")
                
                chinese_name = str(row['品名']) if pd.notna(row['品名']) else ""
                image_path = str(row['图片路径']) if pd.notna(row['图片路径']) else ""
                has_image = bool(row['有图片']) if '有图片' in row else False
                
                # 跳过有图片的产品
                if has_image:
                    print(f"跳过已有图片的产品: {chinese_name}")
                    continue
                
                # 使用映射获取英文名
                english_name = name_map.get(chinese_name, "")
                if not english_name:
                    print(f"警告: 产品 '{chinese_name}' 没有对应的英文名，跳过上传")
                    continue
                
                print(f"正在上传产品: {english_name} (原名: {chinese_name})")
                print(f"产品没有图片，将进行上传")
                current_operation = "准备导航到添加新产品页面"
                
                # 直接导航到添加新产品页面
                print("导航到添加新产品页面...")
                driver.get(f"{wp_url}/wp-admin/post-new.php?post_type=product")
                
                # 确保已经到达添加新产品页面
                try:
                    current_operation = "等待添加新产品页面加载"
                    # 等待页面标题元素加载，确认已经在添加新产品页面
                    WebDriverWait(driver, 15).until(
                        EC.presence_of_element_located((By.ID, "title"))
                    )
                    # 额外检查页面URL
                    current_url = driver.current_url
                    if "post-new.php" in current_url and "post_type=product" in current_url:
                        print("已确认进入添加新产品页面")
                    else:
                        print(f"当前URL: {current_url}，不是添加新产品页面，重新尝试...")
                        driver.get(f"{wp_url}/wp-admin/post-new.php?post_type=product")
                        time.sleep(3)
                        WebDriverWait(driver, 15).until(
                            EC.presence_of_element_located((By.ID, "title"))
                        )
                        if "post-new.php" not in driver.current_url or "post_type=product" not in driver.current_url:
                            raise Exception("无法导航到添加新产品页面")
                        print("已重新导航到添加新产品页面")
                except Exception as page_error:
                    print(f"无法进入添加新产品页面: {page_error}")
                    print(f"当前处理的产品: {chinese_name} ({english_name})")
                    print("跳过当前产品，尝试下一个")
                    continue  # 如果无法进入添加产品页面，直接跳过当前产品
                
                # 确保页面完全加载
                time.sleep(2)
                
                # 移除点击"添加新产品"按钮的部分，因为已经在添加新产品页面了
                
                # 3. 填写产品信息
                print("3. 填写产品信息...")
                current_operation = "填写产品标题"
                # 标题 - 使用英文品名
                title_field = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.ID, "title"))
                )
                title_field.clear()
                # 修改产品标题格式，将连字符"-"改为空格
                product_title = f"{brand} {model} {english_name}"
                title_field.send_keys(product_title)
                print(f"已填写产品标题: {product_title}")
                
                # 跳过描述填写
                print("跳过产品描述填写")
                
                # 4. 设置产品价格 - 直接滚动到常规售价输入框
                print("4. 设置产品价格...")
                current_operation = "滚动到常规售价输入框"
                
                # 尝试直接滚动到常规售价输入框
                try:
                    # 尝试找到价格字段或其标签
                    price_field_or_label = None
                    try:
                        # 先尝试找价格字段
                        price_field_or_label = driver.find_element(By.ID, "_regular_price")
                    except:
                        # 如果找不到价格字段，尝试找标签
                        try:
                            price_field_or_label = driver.find_element(By.XPATH, "//label[contains(text(), '常规售价') or contains(text(), 'Regular price')]")
                        except:
                            # 如果都找不到，尝试找产品数据面板
                            price_field_or_label = driver.find_element(By.ID, "product_data")
                    
                    # 滚动到元素位置
                    if price_field_or_label:
                        driver.execute_script("arguments[0].scrollIntoView({behavior: 'smooth', block: 'center'});", price_field_or_label)
                        print("已滚动页面到价格字段区域")
                        time.sleep(1)  # 等待滚动完成
                except Exception as scroll_error:
                    print(f"滚动页面到价格字段时出错: {scroll_error}")
                    # 尝试通用滚动
                    try:
                        driver.execute_script("window.scrollBy(0, 500);")
                        print("已执行通用页面滚动")
                        time.sleep(1)
                    except:
                        print("无法滚动页面")
                
                # 聚焦并填写价格
                current_operation = "设置产品价格"
                try:
                    # 尝试直接查找价格字段
                    price_field = WebDriverWait(driver, 5).until(
                        EC.presence_of_element_located((By.ID, "_regular_price"))
                    )
                    # 聚焦到价格输入框
                    driver.execute_script("arguments[0].focus();", price_field)
                    price_field.clear()
                    price_field.send_keys(price)
                    print(f"已设置产品价格: {price}")
                except Exception as price_error:
                    print(f"设置产品价格时出错: {price_error}")
                    # 尝试通过JavaScript直接设置价格
                    try:
                        driver.execute_script(f"document.getElementById('_regular_price').value = '{price}';")
                        print(f"通过JavaScript设置产品价格: {price}")
                    except Exception as js_price_error:
                        print(f"通过JavaScript设置产品价格时出错: {js_price_error}")
                        print("无法设置产品价格，但将继续上传产品")
                
                # 5. 处理产品分类
                current_operation = "处理产品分类"
                print("5. 处理产品分类...")
                
                # 等待产品分类面板加载
                try:
                    # 查找产品分类面板
                    WebDriverWait(driver, 10).until(
                        EC.presence_of_element_located((By.ID, "product_catchecklist"))
                    )
                    
                    # 使用映射表中的英文名称作为产品分类
                    category_found = False
                    if english_name:  # 确保英文名不为空
                        print(f"查找产品分类: {english_name}")
                        category_items = driver.find_elements(By.CSS_SELECTOR, "#product_catchecklist li label")
                        
                        for item in category_items:
                            item_text = item.text.strip()
                            if english_name.lower() in item_text.lower():
                                # 找到匹配的英文分类，点击选择
                                print(f"找到匹配的产品分类: {item_text}")
                                checkbox = item.find_element(By.CSS_SELECTOR, "input[type='checkbox']")
                                if not checkbox.is_selected():
                                    driver.execute_script("arguments[0].click();", checkbox)
                                    print(f"已选择产品分类: {item_text}")
                                else:
                                    print(f"产品分类已被选中: {item_text}")
                                category_found = True
                                break
                    
                    # 如果英文分类不存在，则添加新分类
                    if not category_found and english_name:
                        print(f"未找到产品分类: {english_name}，将添加新产品分类")
                        
                        # 点击"添加新分类"链接
                        add_new_cat_toggle = WebDriverWait(driver, 5).until(
                            EC.element_to_be_clickable((By.ID, "product_cat-add-toggle"))
                        )
                        driver.execute_script("arguments[0].click();", add_new_cat_toggle)
                        time.sleep(1)
                        
                        # 输入新英文分类
                        new_cat_input = WebDriverWait(driver, 5).until(
                            EC.presence_of_element_located((By.ID, "newproduct_cat"))
                        )
                        new_cat_input.clear()
                        new_cat_input.send_keys(english_name)
                        
                        # 点击添加按钮
                        add_cat_button = WebDriverWait(driver, 5).until(
                            EC.element_to_be_clickable((By.ID, "product_cat-add-submit"))
                        )
                        driver.execute_script("arguments[0].click();", add_cat_button)
                        
                        # 等待新分类添加完成并被选中
                        time.sleep(2)
                        print(f"已添加并选择新产品分类: {english_name}")
                        
                        # 刷新分类列表，确保新添加的分类被选中
                        category_items = driver.find_elements(By.CSS_SELECTOR, "#product_catchecklist li label")
                        for item in category_items:
                            item_text = item.text.strip()
                            if english_name.lower() in item_text.lower():
                                checkbox = item.find_element(By.CSS_SELECTOR, "input[type='checkbox']")
                                if not checkbox.is_selected():
                                    driver.execute_script("arguments[0].click();", checkbox)
                                break
                except Exception as cat_error:
                    print(f"处理产品分类时出错: {cat_error}")
                    print("继续上传产品，但产品分类可能未正确设置")
                
                # 6. 处理品牌
                current_operation = "处理品牌"
                print("6. 处理品牌...")
                
                # 滚动到品牌选择区域
                try:
                    # 尝试找到品牌面板
                    brand_panel = None
                    try:
                        brand_panel = driver.find_element(By.ID, "product_brandchecklist")
                    except:
                        # 如果找不到，尝试找品牌区域的标题或其他相关元素
                        try:
                            brand_panel = driver.find_element(By.XPATH, "//h2[contains(text(), '品牌') or contains(text(), 'Brand')]")
                        except:
                            print("找不到品牌面板，尝试通用滚动")
                    
                    # 滚动到品牌面板
                    if brand_panel:
                        driver.execute_script("arguments[0].scrollIntoView({behavior: 'smooth', block: 'center'});", brand_panel)
                        print("已滚动页面到品牌选择区域")
                        time.sleep(1)  # 等待滚动完成
                    else:
                        # 如果找不到品牌面板，尝试通用滚动
                        driver.execute_script("window.scrollBy(0, 300);")
                        print("已执行通用页面滚动以寻找品牌区域")
                        time.sleep(1)
                except Exception as brand_scroll_error:
                    print(f"滚动到品牌区域时出错: {brand_scroll_error}")
                    # 尝试通用滚动
                    try:
                        driver.execute_script("window.scrollBy(0, 300);")
                        print("已执行通用页面滚动")
                        time.sleep(1)
                    except:
                        print("无法滚动页面")
                
                # 处理品牌选择
                try:
                    # 查找品牌面板
                    WebDriverWait(driver, 10).until(
                        EC.presence_of_element_located((By.ID, "product_brandchecklist"))
                    )
                    
                    # 检查品牌是否已存在于品牌列表中
                    brand_found = False
                    if brand:  # 确保品牌名不为空
                        print(f"在品牌列表中查找: {brand}")
                        brand_items = driver.find_elements(By.CSS_SELECTOR, "#product_brandchecklist li label")
                        
                        for item in brand_items:
                            item_text = item.text.strip()
                            if brand.lower() in item_text.lower():
                                # 找到匹配的品牌，点击选择
                                print(f"找到匹配的品牌: {item_text}")
                                checkbox = item.find_element(By.CSS_SELECTOR, "input[type='checkbox']")
                                if not checkbox.is_selected():
                                    driver.execute_script("arguments[0].click();", checkbox)
                                    print(f"已选择品牌: {item_text}")
                                else:
                                    print(f"品牌已被选中: {item_text}")
                                brand_found = True
                                break
                    
                    # 如果品牌不存在，则添加新品牌
                    if not brand_found and brand:
                        print(f"未找到品牌: {brand}，将添加新品牌")
                        
                        # 点击"添加新品牌"链接
                        add_new_brand_toggle = WebDriverWait(driver, 5).until(
                            EC.element_to_be_clickable((By.ID, "product_brand-add-toggle"))
                        )
                        driver.execute_script("arguments[0].click();", add_new_brand_toggle)
                        time.sleep(1)
                        
                        # 输入新品牌名称
                        new_brand_input = WebDriverWait(driver, 5).until(
                            EC.presence_of_element_located((By.ID, "newproduct_brand"))
                        )
                        new_brand_input.clear()
                        new_brand_input.send_keys(brand)
                        
                        # 点击添加按钮
                        add_brand_button = WebDriverWait(driver, 5).until(
                            EC.element_to_be_clickable((By.ID, "product_brand-add-submit"))
                        )
                        driver.execute_script("arguments[0].click();", add_brand_button)
                        
                        # 等待新品牌添加完成并被选中
                        time.sleep(2)
                        print(f"已添加并选择新品牌: {brand}")
                        
                        # 刷新品牌列表，确保新添加的品牌被选中
                        brand_items = driver.find_elements(By.CSS_SELECTOR, "#product_brandchecklist li label")
                        for item in brand_items:
                            item_text = item.text.strip()
                            if brand.lower() in item_text.lower():
                                checkbox = item.find_element(By.CSS_SELECTOR, "input[type='checkbox']")
                                if not checkbox.is_selected():
                                    driver.execute_script("arguments[0].click();", checkbox)
                                break
                except Exception as brand_error:
                    print(f"处理品牌时出错: {brand_error}")
                    print("继续上传产品，但品牌可能未正确设置")
                
                # 7. 跳过产品图片上传（因为我们只处理没有图片的产品）
                print("7. 跳过产品图片上传（产品没有图片）...")
                current_operation = "跳过产品图片上传"
                
                # 8. 发布产品前的最终检查
                print("8. 发布产品前的最终检查...")
                current_operation = "发布产品前的最终检查"
                
                # 检查产品标题是否已填写
                title_value = driver.find_element(By.ID, "title").get_attribute("value")
                if not title_value:
                    print("警告: 产品标题为空，尝试重新填写")
                    title_field = driver.find_element(By.ID, "title")
                    title_field.clear()
                    product_title = f"{brand} {model} {english_name}"
                    title_field.send_keys(product_title)
                
                # 检查产品价格是否已填写
                try:
                    price_field = driver.find_element(By.ID, "_regular_price")
                    price_value = price_field.get_attribute("value")
                    if not price_value and price:
                        print("警告: 产品价格为空，尝试重新填写")
                        price_field.clear()
                        price_field.send_keys(price)
                except:
                    print("警告: 无法检查产品价格")
                
                # 检查产品分类是否已选择
                try:
                    if english_name:
                        category_selected = False
                        category_items = driver.find_elements(By.CSS_SELECTOR, "#product_catchecklist li input:checked")
                        if len(category_items) > 0:
                            category_selected = True
                        
                        if not category_selected:
                            print("警告: 产品分类未选择，尝试重新选择")
                            # 尝试再次选择产品分类
                            category_items = driver.find_elements(By.CSS_SELECTOR, "#product_catchecklist li label")
                            for item in category_items:
                                item_text = item.text.strip()
                                if english_name.lower() in item_text.lower():
                                    checkbox = item.find_element(By.CSS_SELECTOR, "input[type='checkbox']")
                                    if not checkbox.is_selected():
                                        driver.execute_script("arguments[0].click();", checkbox)
                                    break
                except Exception as check_error:
                    print(f"检查产品分类时出错: {check_error}")
                
                # 9. 发布产品
                print("9. 发布没有图片的产品...")
                current_operation = "等待发布按钮变为可点击状态"
                
                # 等待发布按钮变为可点击状态
                try:
                    print("等待发布按钮变为可点击状态...")
                    # 首先找到发布按钮，无论其状态如何
                    publish_button = WebDriverWait(driver, 5).until(
                        EC.presence_of_element_located((By.ID, "publish"))
                    )
                    
                    # 滚动到发布按钮，确保它在视图中
                    driver.execute_script("arguments[0].scrollIntoView({behavior: 'smooth', block: 'center'});", publish_button)
                    time.sleep(1)  # 等待滚动完成
                    
                    # 等待按钮变为可点击状态（最多等待30秒）
                    max_wait_time = 30
                    wait_interval = 1
                    total_waited = 0
                    
                    while total_waited < max_wait_time:
                        # 检查按钮是否可点击
                        is_disabled = driver.execute_script(
                            "return arguments[0].disabled === true || arguments[0].classList.contains('disabled') || arguments[0].getAttribute('aria-disabled') === 'true';", 
                            publish_button
                        )
                        
                        if not is_disabled:
                            print(f"发布按钮已变为可点击状态，等待了{total_waited}秒")
                            break
                        
                        print(f"发布按钮仍处于不可点击状态，已等待{total_waited}秒...")
                        time.sleep(wait_interval)
                        total_waited += wait_interval
                    
                    if total_waited >= max_wait_time:
                        print("警告：发布按钮在最大等待时间内未变为可点击状态，将尝试点击")
                    
                    # 再次等待一小段时间，确保按钮完全可用
                    time.sleep(2)
                    
                    # 现在尝试点击发布按钮
                    current_operation = "点击发布按钮"
                    
                    # 确保发布按钮可点击
                    publish_button = WebDriverWait(driver, 10).until(
                        EC.element_to_be_clickable((By.ID, "publish"))
                    )
                    
                    # 尝试多种方式点击发布按钮
                    try:
                        # 方法1：直接点击
                        publish_button.click()
                        print("方法1：直接点击发布按钮")
                    except Exception as click_error:
                        print(f"直接点击发布按钮失败: {click_error}")
                        try:
                            # 方法2：使用JavaScript点击
                            driver.execute_script("arguments[0].click();", publish_button)
                            print("方法2：使用JavaScript点击发布按钮")
                        except Exception as js_click_error:
                            print(f"JavaScript点击发布按钮失败: {js_click_error}")
                            try:
                                # 方法3：使用ActionChains点击
                                from selenium.webdriver.common.action_chains import ActionChains
                                actions = ActionChains(driver)
                                actions.move_to_element(publish_button).click().perform()
                                print("方法3：使用ActionChains点击发布按钮")
                            except Exception as action_click_error:
                                print(f"ActionChains点击发布按钮失败: {action_click_error}")
                                # 方法4：尝试查找所有可能的发布按钮并点击
                                publish_buttons = driver.find_elements(By.XPATH, "//input[@id='publish' or @name='publish' or @value='发布' or @value='Publish']")
                                if publish_buttons:
                                    driver.execute_script("arguments[0].click();", publish_buttons[0])
                                    print("方法4：找到并点击备选发布按钮")
                                else:
                                    raise Exception("无法找到任何可用的发布按钮")
                    
                    print("已尝试点击发布按钮")
                    
                    # 等待发布完成 - 检测成功消息或新页面加载
                    try:
                        WebDriverWait(driver, 15).until(
                            EC.presence_of_element_located((By.CLASS_NAME, "updated"))
                        )
                        print("检测到发布成功消息")
                    except:
                        # 如果没有找到成功消息，检查是否已重定向到新页面
                        try:
                            WebDriverWait(driver, 15).until(
                                EC.presence_of_element_located((By.ID, "title"))
                            )
                            print("已重定向到新页面，发布可能成功")
                        except:
                            # 如果上述两种方法都失败，检查URL是否已更改
                            if "post.php" in driver.current_url and "post_type=product" in driver.current_url:
                                print("URL已更改为编辑页面，发布可能成功")
                            else:
                                print("警告：无法确认发布是否成功，但将继续处理")
                    
                    print(f"没有图片的产品已尝试上传: {english_name}")
                    upload_count += 1
                    time.sleep(2)  # 防止请求过快
                except Exception as publish_error:
                    print(f"发布产品时出错: {publish_error}")
                    # 尝试再次点击发布按钮
                    try:
                        publish_buttons = driver.find_elements(By.XPATH, "//input[@id='publish' or @name='publish' or @value='发布' or @value='Publish']")
                        if publish_buttons:
                            driver.execute_script("arguments[0].click();", publish_buttons[0])
                            time.sleep(5)
                            print("通过备选方法点击发布按钮")
                            upload_count += 1
                        else:
                            print("找不到发布按钮，尝试通过键盘快捷键发布")
                            # 尝试使用键盘快捷键 Ctrl+S 发布
                            from selenium.webdriver.common.keys import Keys
                            from selenium.webdriver.common.action_chains import ActionChains
                            actions = ActionChains(driver)
                            actions.key_down(Keys.CONTROL).send_keys('s').key_up(Keys.CONTROL).perform()
                            time.sleep(5)
                            print("通过键盘快捷键尝试发布")
                            upload_count += 1
                    except Exception as alt_publish_error:
                        print(f"备选发布方法也失败: {alt_publish_error}")
                        print("无法发布产品，跳过当前产品")
                
                # 为下一个产品直接导航到添加新产品页面
                print("导航到添加新产品页面准备上传下一个产品...")
                driver.get(f"{wp_url}/wp-admin/post-new.php?post_type=product")
                WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.ID, "title"))
                )
                time.sleep(1)
            except Exception as product_error:
                print(f"处理产品时出错: {product_error}")
                print(f"出错时正在处理的产品: {chinese_name} ({english_name if 'english_name' in locals() else '未获取英文名'})")
                print(f"出错时正在执行的操作: {current_operation if 'current_operation' in locals() else '未知操作'}")
                print("跳过当前产品，继续下一个")
                # 确保即使出错也能回到添加产品页面
                try:
                    driver.get(f"{wp_url}/wp-admin/post-new.php?post_type=product")
                    WebDriverWait(driver, 10).until(
                        EC.presence_of_element_located((By.ID, "title"))
                    )
                    time.sleep(1)
                except:
                    print("无法导航回添加产品页面，尝试继续...")
                continue
        
        print(f"成功上传 {upload_count} 个产品")
    except Exception as e:
        print(f"上传过程中出错: {e}")
        # 不要在这里尝试访问可能未定义的变量
    finally:
        driver.quit()

def main():
    excel_file = "a.xlsx"
    
    # 读取Excel文件
    print("\n= 步骤1: 读取Excel文件 =")
    df = read_excel(excel_file)
    if df is None or len(df) == 0:
        print("Excel文件为空或读取失败，程序结束")
        return
    
    # 准备产品数据（替换原来的检查图片步骤）
    print("\n= 步骤2: 准备产品数据 =")
    df = prepare_product_data(df)
    
    # 创建中英文品名映射表
    print("\n= 步骤3: 创建中英文品名映射表 =")
    mapping_file = "name_mapping_new.xlsx"
    
    # 检查映射文件是否已存在
    if os.path.exists(mapping_file):
        print(f"发现已存在的映射表: {mapping_file}")
        use_existing = input("是否使用已存在的映射表? (y/n): ")
        if use_existing.lower() == 'y':
            print(f"将使用已存在的映射表: {mapping_file}")
        else:
            # 用户选择创建新的映射表
            mapping_file = create_name_mapping(df, mapping_file)
            if not mapping_file:
                print("创建映射表失败，无法继续")
                return
    else:
        # 映射文件不存在，创建新的
        mapping_file = create_name_mapping(df, mapping_file)
        if not mapping_file:
            print("创建映射表失败，无法继续")
            return
    
    # 在此提示用户填写映射表
    check_mapping = input(f"请确认 {mapping_file} 中已填写好英文品名，是否继续? (y/n): ")
    if check_mapping.lower() != 'y':
        print("已取消操作")
        return
    
    # 读取映射表
    name_map = read_mapping(mapping_file)
    if not name_map:
        print("映射表为空或读取失败，无法继续")
        return
    
    # 询问WordPress登录信息
    print("\n= 步骤4: 上传产品到WordPress =")
    wp_url = input("请输入WordPress网站地址 (例如: https://example.com): ")
    username = input("请输入WordPress用户名: ")
    password = input("请输入WordPress密码: ")
    
    # 确认上传
    confirm = input(f"将上传 {len(df)} 个产品到 {wp_url}，确认继续? (y/n): ")
    if confirm.lower() != 'y':
        print("已取消上传")
        return
    
    # 上传产品
    upload_to_wordpress(df, name_map, wp_url, username, password)
    
    print("所有操作已完成")

if __name__ == "__main__":
    main()