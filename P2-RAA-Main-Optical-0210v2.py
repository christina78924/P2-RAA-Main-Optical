import pandas as pd
from pptx import Presentation
from pptx.util import Inches
import glob
import os
import subprocess
import time
import warnings

# 忽略 Pandas 的警告訊息
warnings.filterwarnings("ignore")

# ==========================================
# 1. 設定區域
# ==========================================

# ⚠️ 請確認您的 Mac JMP 應用程式名稱
# 通常是 "JMP 17" 或 "JMP Pro 17"，如果不確定，可以先試試看 "JMP 17"
JMP_APP_NAME = "JMP 19" 

# 設定暫存檔案名稱
TEMP_DATA_CSV = "temp_jmp_data.csv"
TEMP_JSL_FILE = "temp_plot.jsl"

# ==========================================
# 2. 數據處理函數
# ==========================================

def find_header_row(file_path, sheet_name):
    """
    自動尋找 Excel 中 'Tester' 開頭的那一行作為標題列
    """
    try:
        # 先讀取前 20 行
        df_temp = pd.read_excel(file_path, sheet_name=sheet_name, header=None, nrows=20, engine='openpyxl')
        for idx, row in df_temp.iterrows():
            if isinstance(row[0], str) and row[0].strip().startswith('Tester'):
                return idx
        return 0 
    except Exception as e:
        print(f"⚠️ 無法讀取 Header ({sheet_name}): {e}")
        return 0

def get_station_name(col_name):
    """
    將欄位名稱簡化為標準站點名稱
    """
    if 'PreAA' in col_name:
        if 'H1' in col_name or 'V1' in col_name: return 'PreAA_1'
        if 'H2' in col_name or 'V2' in col_name: return 'PreAA_2'
    if 'AfterExposure' in col_name: return 'AfterExp'
    if 'LooseClaws' in col_name: return 'LooseClaws'
    if 'AA_M87' in col_name: return 'AA'
    if 'AfterBaking' in col_name: return 'AfterBaking'
    return None

def process_data():
    """
    讀取資料夾內所有 Excel，合併並清洗數據
    """
    all_data = []
    excel_files = glob.glob('*.xlsx')
    
    if not excel_files:
        print("❌ 找不到 Excel 檔案！")
        return pd.DataFrame()

    print(f"📂 找到 {len(excel_files)} 個 Excel 檔案，開始處理...")

    for file in excel_files:
        try:
            xls = pd.ExcelFile(file, engine='openpyxl')
        except:
            continue

        for sheet in xls.sheet_names:
            if sheet not in ['RAA-R', 'RAA-L', 'IPQC-R', 'IPQC-L']:
                continue
            
            print(f"  -> 讀取: {file} [{sheet}]")
            header_idx = find_header_row(file, sheet)
            df = pd.read_excel(file, sheet_name=sheet, header=header_idx, engine='openpyxl')
            
            side = 'Right' if '-R' in sheet else 'Left'
            
            # 抓取關鍵欄位
            target_cols = [c for c in df.columns if 'Boresight' in str(c) and 'White' in str(c)]
            if not target_cols: continue

            if 'CreateTime' in df.columns:
                df['CreateTime'] = pd.to_datetime(df['CreateTime'], errors='coerce')
            
            # 轉置數據 (Melt)
            melted = df.melt(id_vars=['CreateTime'], value_vars=target_cols, 
                             var_name='Station_Raw', value_name='Value')
            melted['Side'] = side
            
            # 判斷 H/V 方向與站點
            def get_direction(name):
                if '_H_' in name or 'illu_Boresight_H' in name: return 'H'
                if '_V_' in name or 'illu_Boresight_V' in name: return 'V'
                return 'Unknown'
            
            melted['Direction'] = melted['Station_Raw'].apply(get_direction)
            melted['Station_Generic'] = melted['Station_Raw'].apply(get_station_name)
            
            all_data.append(melted)

    if not all_data: return pd.DataFrame()

    final_df = pd.concat(all_data, ignore_index=True)
    final_df['Value'] = pd.to_numeric(final_df['Value'], errors='coerce')
    final_df = final_df.dropna(subset=['Value', 'Station_Generic'])
    
    # 建立顯示用的標籤 (讓 L 在左，R 在右)
    final_df['Display_Label'] = final_df['Side'].str[0] + '-' + final_df['Station_Generic']
    
    # 建立排序索引 (為了讓 JMP 圖表依正確順序排列)
    station_order = ['PreAA_1', 'PreAA_2', 'AA', 'AfterExp', 'LooseClaws', 'AfterBaking']
    order_map = {name: i for i, name in enumerate(station_order)}
    
    def get_sort_key(row):
        base_order = order_map.get(row['Station_Generic'], 99)
        # Left = 0~99, Right = 100~199
        return base_order if row['Side'] == 'Left' else base_order + 100

    final_df['Sort_Key'] = final_df.apply(get_sort_key, axis=1)
    
    return final_df

# ==========================================
# 3. JMP 繪圖控制核心 (Mac 版)
# ==========================================

def run_jmp_on_mac(df, chart_type, output_image_name, ylim=None):
    """
    生成 JSL -> 呼叫 JMP -> 等待產圖
    """
    abs_csv_path = os.path.abspath(TEMP_DATA_CSV)
    abs_img_path = os.path.abspath(output_image_name)
    
    # 1. 儲存數據給 JMP 用
    df.to_csv(abs_csv_path, index=False)
    
    # 若舊圖存在，先刪除，以免誤判
    if os.path.exists(abs_img_path):
        os.remove(abs_img_path)

    # 2. 準備 JSL 腳本內容
    jsl_content = ""
    
    if chart_type == 'boxplot':
        # Box Plot JSL
        # 設定 Y 軸範圍字串
        scale_script = ""
        if ylim:
            scale_script = f'Min( {ylim[0]} ), Max( {ylim[1]} ),'
            
        jsl_content = f"""
        Names Default To Here( 1 );
        dt = Open( "{abs_csv_path}" );
        
        // 確保依照 Sort_Key 排序 X 軸
        dt << Sort( By( :Sort_Key ), Order( Ascending ), Replace Table );

        gb = dt << Graph Builder(
            Size( 1000, 800 ),
            Show Control Panel( 0 ),
            Variables( X( :Display_Label ), Y( :Value ), Group X( :Side ), Group Y( :Direction ) ),
            Elements( Box Plot( X, Y, Legend( 5 ) ) ),
            SendToReport(
                Dispatch( {{}}, "Value", ScaleBox, 
                    {{ {scale_script} Add Ref Line( 0.25, "Solid", "Red", "USL", 2 ), 
                      Add Ref Line( -0.25, "Solid", "Red", "LSL", 2 ) }} 
                )
            )
        );
        
        gb << Save Picture( "{abs_img_path}", "PNG" );
        Close( dt, NoSave );
        Exit(); // 執行完後關閉 JMP (如果不希望關閉，請拿掉這行)
        """
        
    elif chart_type == 'control_chart':
        # Control Chart JSL
        scale_script = ""
        if ylim:
            scale_script = f'Min( {ylim[0]} ), Max( {ylim[1]} ),'

        jsl_content = f"""
        Names Default To Here( 1 );
        dt = Open( "{abs_csv_path}" );
        
        // 只取 AfterBaking
        dt << Select Where( :Station_Generic == "AfterBaking" );
        dt_sub = dt << Subset( Selected Rows( 1 ), Output Table( "Sub" ) );
        Close( dt, NoSave );
        
        gb = dt_sub << Graph Builder(
            Size( 1200, 800 ),
            Show Control Panel( 0 ),
            Variables( X( :CreateTime ), Y( :Value ), Group X( :Side ), Group Y( :Direction ) ),
            Elements( Points( X, Y, Legend( 3 ) ) ),
            SendToReport(
                Dispatch( {{}}, "Value", ScaleBox, 
                    {{ {scale_script} Add Ref Line( 0.25, "Solid", "Red", "USL", 2 ), 
                      Add Ref Line( -0.25, "Solid", "Red", "LSL", 2 ) }} 
                )
            )
        );
        
        gb << Save Picture( "{abs_img_path}", "PNG" );
        Close( dt_sub, NoSave );
        Exit();
        """

    # 3. 寫入 JSL 檔案
    with open(TEMP_JSL_FILE, 'w', encoding='utf-8') as f:
        f.write(jsl_content)
        
    # 4. 呼叫 Mac JMP 執行
    print(f"   -> 正在呼叫 JMP 繪製: {output_image_name} ...")
    try:
        # 使用 'open' 指令
        subprocess.run(['open', '-a', JMP_APP_NAME, TEMP_J