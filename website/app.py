from flask import Flask, request, render_template, send_file, jsonify
import os
import pandas as pd
import json
import io
from processing import supabase
from processing import main
from formatting import style_excel

app = Flask(__name__)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

#For temporary Storage
DAILY_UPLOAD_FOLDER = os.path.join(BASE_DIR, 'Daily_Uploads')
PROCESSED_FOLDER = os.path.join(BASE_DIR, 'Processed')

os.makedirs(DAILY_UPLOAD_FOLDER, exist_ok=True)
os.makedirs(PROCESSED_FOLDER, exist_ok=True)


MAX_FILE_SIZE = 1.5 * 1024 * 1024


@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        if 'file' not in request.files:
            return 'Error: no file uploaded', 400
        
        uploaded_file = request.files['file']
        if not uploaded_file or uploaded_file.filename == '':
            return 'Error: no file selected', 400

        uploaded_file.seek(0, 2)
        file_size = uploaded_file.tell()
        uploaded_file.seek(0)
        
        if file_size > MAX_FILE_SIZE:
            size_mb = file_size / (1024 * 1024)
            return f'Error: File size ({size_mb:.2f}MB) exceeds the maximum allowed size of 1.5MB', 400

        try:
            temp_name = f"temp_{uploaded_file.filename}"
            temp_path = os.path.join(DAILY_UPLOAD_FOLDER, temp_name)
            os.makedirs(DAILY_UPLOAD_FOLDER, exist_ok=True)
            uploaded_file.save(temp_path)

            print('Beginning processing...')
            results = main(temp_path)

            date_of_report = results['date_of_report']
            year_of_report = results['year_of_report']
            
            prefix = date_of_report[:4]
            
            if prefix.isdigit():
                report_date_str = f"{date_of_report}"
                print('formatting .xlsx as a date')
            else:
                report_date_str = f"{date_of_report} {year_of_report}"
                print('formatting .xlsx as a range')
                
            ext = os.path.splitext(uploaded_file.filename)[1]
            standard_name = f"{report_date_str} {ext}"
            standard_path = os.path.join(DAILY_UPLOAD_FOLDER, standard_name)
            
            if os.path.exists(standard_path):
                print(f"Removing old file: {standard_name}")
                os.remove(standard_path)
            
            os.rename(temp_path, standard_path)
            print(f"Saved as: {standard_name}")

            base_out_name = f"PLB Tabulation {report_date_str}.xlsx"
            output_file_path = os.path.join(PROCESSED_FOLDER, base_out_name)
            
            if os.path.exists(output_file_path):
                print(f"Removing old output file: {base_out_name}")
                os.remove(output_file_path)
            
            results['output_file'] = output_file_path
            
            print('Beginning Report Formatting...')
            output_file = style_excel(results)

            print('Report Generated!')
            print(f'Uploading {base_out_name} to Supabase (will overwrite if exists)...')
            upload_to_supabase(output_file, base_out_name)

            return send_file(output_file, as_attachment=True, download_name=base_out_name)
        
        except Exception as e:
            print(f"Error in daily upload: {e}")
            import traceback
            traceback.print_exc()
            return f'Error processing daily file: {str(e)}', 500
        
    return render_template('index.html')

@app.route('/api/commandos_namelist', methods=['GET'])
def get_commandos_namelist():
    try:
        response = supabase.table("commando_namelist").select("name").execute()
        names = [row['name'] for row in response.data]
        return jsonify({'names': names, 'success': True})
    except Exception as e:
        print(f"Error fetching commandos namelist: {e}")
        return jsonify({'names': [], 'error': str(e), 'success': False}), 500

@app.route('/api/st_namelist', methods=['GET'])
def get_st_namelist():
    try:
        response = supabase.table("st_namelist").select("name").execute()
        names = [row['name'] for row in response.data]
        return jsonify({'names': names, 'success': True})
    except Exception as e:
        print(f"Error fetching ST namelist: {e}")
        return jsonify({'names': [], 'error': str(e), 'success': False}), 500

@app.route('/api/commandos_namelist/delete', methods=['POST'])
def delete_commandos_namelist():
    try:
        data = request.get_json()
        name_to_delete = data.get('name')
        if not name_to_delete:
            return jsonify({'error': 'No name provided', 'success': False}), 400

        response = supabase.table("commando_namelist") \
            .delete() \
            .eq("name", name_to_delete) \
            .execute()

        if not response.data:
            return jsonify({'error': f'Name "{name_to_delete}" not found in database', 'success': False}), 404

        return jsonify({'success': True, 'deleted': response.data})
    except Exception as e:
        print(f"Error deleting from commandos namelist: {e}")
        return jsonify({'error': str(e), 'success': False}), 500

@app.route('/api/st_namelist/delete', methods=['POST'])
def delete_st_namelist():
    try:
        data = request.get_json()
        name_to_delete = data.get('name')
        if not name_to_delete:
            return jsonify({'error': 'No name provided', 'success': False}), 400

        response = supabase.table("st_namelist") \
            .delete() \
            .eq("name", name_to_delete) \
            .execute()

        if not response.data:
            return jsonify({'error': f'Name "{name_to_delete}" not found in ST List', 'success': False}), 404

        return jsonify({'success': True})
    except Exception as e:
        print(f"Error deleting from ST namelist: {e}")
        return jsonify({'error': str(e), 'success': False}), 500

@app.route('/api/add-name', methods=['POST'])
def add_name():
    try:
        data = request.json
        name = data['name'].strip()
        team = data['team']

        table_name = "commando_namelist" if team == 'commandos' else "st_namelist"

        check_res = supabase.table(table_name) \
            .select("name") \
            .eq("name", name) \
            .execute()

        if not check_res.data:
            supabase.table(table_name) \
                .insert({"name": name}) \
                .execute()
            return jsonify({'message': f'Added to {team}', 'success': True})
        else:
            return jsonify({'message': f'Already in {team}', 'success': True})
    except Exception as e:
        print(f"Error adding name: {e}")
        return jsonify({'error': str(e), 'success': False}), 500

@app.route('/api/check-new-names', methods=['POST'])
def check_new_names():
    try:
        uploaded_file = request.files['file']
        
        uploaded_file.seek(0, 2)
        file_size = uploaded_file.tell()
        uploaded_file.seek(0)
        
        if file_size > MAX_FILE_SIZE:
            size_mb = file_size / (1024 * 1024)
            return jsonify({
                'error': f'File size ({size_mb:.2f}MB) exceeds maximum of 1.5MB',
                'success': False
            }), 400
        
        df = pd.read_csv(uploaded_file) if uploaded_file.filename.endswith('.csv') else pd.read_excel(uploaded_file)

        commando_res = supabase.table("commando_namelist").select("name").execute()
        st_res = supabase.table("st_namelist").select("name").execute()

        combined_namelist = set(
            [row['name'] for row in commando_res.data] +
            [row['name'] for row in st_res.data]
        )

        columns_to_check = ['MAB_BY', 'PLB Docking By', 'PLB (R) Arrival By', 'Pax step docking by']
        new_names_found = []

        for i in df.index:
            for col in columns_to_check:
                if col not in df.columns:
                    continue
                name = df.loc[i, col]
                if pd.isna(name):
                    continue
                name = str(name).strip()

                if name and name not in combined_namelist:
                    if not any(n['name'] == name for n in new_names_found):
                        new_names_found.append({'name': name, 'row': i + 2})

        return jsonify({'new_names': new_names_found, 'success': True})
    except Exception as e:
        print(f"Error checking new names: {e}")
        import traceback
        traceback.print_exc()
        return jsonify({'error': str(e), 'success': False}), 500

@app.route('/api/check-new-names-daily', methods=['POST'])
def check_new_names_daily():
    """Same as check_new_names but for daily upload flow"""
    try:
        uploaded_file = request.files['file']
        
        # Check file size
        uploaded_file.seek(0, 2)
        file_size = uploaded_file.tell()
        uploaded_file.seek(0)
        
        if file_size > MAX_FILE_SIZE:
            size_mb = file_size / (1024 * 1024)
            return jsonify({
                'error': f'File size ({size_mb:.2f}MB) exceeds maximum of 1.5MB',
                'success': False
            }), 400

        df = pd.read_csv(uploaded_file) if uploaded_file.filename.endswith('.csv') else pd.read_excel(uploaded_file)

        print(f"Available columns in uploaded file: {df.columns.tolist()}")

        commando_res = supabase.table("commando_namelist").select("name").execute()
        st_res = supabase.table("st_namelist").select("name").execute()

        combined_namelist = set(
            str(row['name']).strip().lower() 
            for row in (commando_res.data + st_res.data)
        )
        
        print(f"Total known names in database: {len(combined_namelist)}")

        columns_to_check = ['MAB_BY', 'PLB Docking By', 'PLB (R) Arrival By', 'Pax step docking by']
        new_names_found = []

        for i in df.index:
            for col in columns_to_check:
                if col not in df.columns:
                    print(f"Column '{col}' not found in uploaded file")
                    continue
                    
                name = df.loc[i, col]
                if pd.isna(name):
                    continue
                    
                name_original = str(name).strip()
                name_normalized = name_original.lower()

                if name_original and name_normalized not in combined_namelist:
                    if not any(n['name'] == name_original for n in new_names_found):
                        new_names_found.append({'name': name_original, 'row': i + 2})
                        print(f"Found new name: '{name_original}' at row {i + 2} in column '{col}'")

        print(f"Total new names found: {len(new_names_found)}")
        return jsonify({'new_names': new_names_found, 'success': True})
        
    except Exception as e:
        print(f"Error checking new names (daily): {e}")
        import traceback
        traceback.print_exc()
        return jsonify({'error': str(e), 'success': False}), 500

@app.route('/api/check-duplicate-dates', methods=['POST'])
def check_duplicate_dates():
    try:
        uploaded_file = request.files['file']
        
        # Check file size
        uploaded_file.seek(0, 2)
        file_size = uploaded_file.tell()
        uploaded_file.seek(0)
        
        if file_size > MAX_FILE_SIZE:
            size_mb = file_size / (1024 * 1024)
            return jsonify({
                'error': f'File size ({size_mb:.2f}MB) exceeds maximum of 1.5MB',
                'success': False
            }), 400
        
        df = pd.read_csv(uploaded_file) if uploaded_file.filename.endswith('.csv') else pd.read_excel(uploaded_file)

        if 'flightDate' not in df.columns:
            return jsonify({
                'error': 'flightDate column not found in uploaded file',
                'available_columns': df.columns.tolist(),
                'success': False
            }), 400

        uploaded_dates = set()
        for date_val in df['flightDate'].dropna():
            try:
                date_str = pd.to_datetime(date_val).strftime('%Y-%m-%d')
                uploaded_dates.add(date_str)
            except Exception as e:
                print(f"Warning: Could not parse date '{date_val}': {e}")

        print(f"\n=== DUPLICATE DATE CHECK ===")
        print(f"Unique dates in uploaded file: {sorted(uploaded_dates)}")

        tables_to_check = [
            "daily_commando",
            "daily_st", 
            "daily_percentage_docked",
            "commando_pivot_data",
            "st_pivot_data",
            "bay_alphabet_data"
        ]

        existing_dates = {}
        for table in tables_to_check:
            try:
                res = supabase.table(table).select("*").execute()
                
                if not res.data:
                    print(f"✓ Table '{table}' is empty")
                    existing_dates[table] = set()
                    continue
                
                date_column = 'date'
                
                table_dates = set()
                for row in res.data:
                    if date_column in row and row[date_column]:
                        try:
                            date_val = row[date_column]

                            if isinstance(date_val, str):
                                date_str = date_val.split('T')[0].split(' ')[0]
                            else:
                                date_str = pd.to_datetime(date_val).strftime('%Y-%m-%d')
                            table_dates.add(date_str)
                        except Exception as e:
                            print(f"Warning: Could not parse date '{row[date_column]}' from {table}: {e}")
                
                existing_dates[table] = table_dates
                print(f"✓ Table '{table}': {len(table_dates)} unique dates")
                if table_dates:
                    print(f"  Sample dates: {sorted(list(table_dates))[:3]}")
                
            except Exception as e:
                print(f"✗ Error fetching dates from {table}: {e}")
                import traceback
                traceback.print_exc()
                existing_dates[table] = set()

        duplicates_by_date = {}
        for date in uploaded_dates:
            tables_with_date = []
            for table, table_dates in existing_dates.items():
                if date in table_dates:
                    tables_with_date.append(table)
            
            if tables_with_date:
                duplicates_by_date[date] = tables_with_date

        print(f"\n=== RESULTS ===")
        if duplicates_by_date:
            print(f"✗ Found {len(duplicates_by_date)} duplicate date(s)")
            for date, tables in duplicates_by_date.items():
                print(f"  {date}: {', '.join(tables)}")
            
            return jsonify({
                'has_duplicates': True,
                'duplicates': duplicates_by_date,
                'total_uploaded_dates': len(uploaded_dates),
                'total_duplicate_dates': len(duplicates_by_date),
                'message': f'Found {len(duplicates_by_date)} date(s) that already exist in the database',
                'success': True
            })
        else:
            print(f"✓ No duplicate dates found - safe to proceed")
            return jsonify({
                'has_duplicates': False,
                'total_uploaded_dates': len(uploaded_dates),
                'message': 'No duplicate dates found - safe to proceed',
                'success': True
            })

    except Exception as e:
        print(f"✗ Error checking duplicate dates: {e}")
        import traceback
        traceback.print_exc()
        return jsonify({'error': str(e), 'success': False}), 500

@app.route('/api/process-with-action', methods=['POST'])
def process_with_action():
    """
    Process uploaded file with specified action for duplicate dates.
    Action can be: 'overwrite' or 'skip'
    """
    try:
        if 'file' not in request.files:
            return 'Error: no file uploaded', 400
        
        uploaded_file = request.files['file']
        if not uploaded_file or uploaded_file.filename == '':
            return 'Error: no file selected', 400
        
        uploaded_file.seek(0, 2)
        file_size = uploaded_file.tell()
        uploaded_file.seek(0)
        
        if file_size > MAX_FILE_SIZE:
            size_mb = file_size / (1024 * 1024)
            return f'Error: File size ({size_mb:.2f}MB) exceeds the maximum allowed size of 1.5MB', 400

        action = request.form.get('action', 'overwrite')
        
        if action not in ['overwrite', 'skip']:
            return 'Error: Invalid action. Must be overwrite or skip', 400

        if action == 'skip':
            return jsonify({'success': True, 'message': 'Upload skipped by user'}), 200

        temp_name = f"temp_{uploaded_file.filename}"
        temp_path = os.path.join(DAILY_UPLOAD_FOLDER, temp_name)
        os.makedirs(DAILY_UPLOAD_FOLDER, exist_ok=True)
        uploaded_file.save(temp_path)

        print(f'Beginning processing with action: {action}...')
        df_temp = pd.read_excel(temp_path, engine='openpyxl')
        dates_in_file = set()
        for date_val in df_temp['flightDate'].dropna():
            try:
                date_str = pd.to_datetime(date_val).strftime('%Y-%m-%d')
                dates_in_file.add(date_str)
            except:
                pass
        
        print(f'Dates in uploaded file: {sorted(dates_in_file)}')
        
        print(f'\n=== OVERWRITE MODE: Deleting existing data for dates: {sorted(dates_in_file)} ===')
        tables_to_delete = [
            "daily_commando",
            "daily_st",
            "daily_percentage_docked",
            "commando_pivot_data",
            "st_pivot_data",
            "bay_alphabet_data"
        ]
        
        for table in tables_to_delete:
            try:
                deleted_count = 0
                for date_to_delete in dates_in_file:
                    delete_res = supabase.table(table).delete().eq("date", date_to_delete).execute()
                    if delete_res.data:
                        deleted_count += len(delete_res.data)
                print(f'✓ Deleted {deleted_count} records from {table}')
            except Exception as e:
                print(f'✗ Error deleting from {table}: {e}')
        
        results = main(temp_path)
        date_of_report = results['date_of_report']
        year_of_report = results['year_of_report']
        
        prefix = date_of_report[:4]
        
        if prefix.isdigit():
            report_date_str = f"{date_of_report}"
            print('formatting .xlsx as a date')
        else:
            report_date_str = f"{date_of_report} {year_of_report}"
            print('formatting .xlsx as a range')

        ext = os.path.splitext(uploaded_file.filename)[1]
        standard_name = f"{report_date_str} {ext}"
        standard_path = os.path.join(DAILY_UPLOAD_FOLDER, standard_name)
        
        if os.path.exists(standard_path):
            print(f"Removing old file: {standard_name}")
            os.remove(standard_path)
        
        os.rename(temp_path, standard_path)

        base_out_name = f"PLB Tabulation {report_date_str}.xlsx"
        output_file_path = os.path.join(PROCESSED_FOLDER, base_out_name)
        
        if os.path.exists(output_file_path):
            print(f"Removing old output file: {base_out_name}")
            os.remove(output_file_path)
        
        results['output_file'] = output_file_path
        
        print('Beginning Report Formatting...')
        output_file = style_excel(results)

        print('Report Generated!')
        print(f'Uploading {base_out_name} to Supabase (will overwrite if exists)...')
        upload_to_supabase(output_file, base_out_name)

        return send_file(output_file, as_attachment=True, download_name=base_out_name)
    
    except Exception as e:
        print(f"Error in process-with-action: {e}")
        import traceback
        traceback.print_exc()
        return f'Error processing file: {str(e)}', 500


@app.route('/api/weekly_data', methods=['GET'])
def get_weekly_data():
    """
    Get all unique dates from all 6 tables, consolidate and return as a single list
    """
    try:
        print("\n=== FETCHING WEEKLY DATA ===")
        
        tables_to_check = [
            "daily_commando",
            "daily_st",
            "daily_percentage_docked",
            "commando_pivot_data",
            "st_pivot_data",
            "bay_alphabet_data"
        ]
        
        all_dates = set()
        date_table_mapping = {}
        
        for table in tables_to_check:
            try:
                res = supabase.table(table).select("*").execute()
                
                if not res.data:
                    print(f"✓ Table '{table}' is empty")
                    continue
                
                date_column = 'date'
                
                for row in res.data:
                    if date_column in row and row[date_column]:
                        try:
                            date_val = row[date_column]
                            
                            if isinstance(date_val, str):
                                date_str = date_val.split('T')[0].split(' ')[0]
                            else:
                                date_str = pd.to_datetime(date_val).strftime('%Y-%m-%d')
                            
                            all_dates.add(date_str)
                            
                            if date_str not in date_table_mapping:
                                date_table_mapping[date_str] = []
                            date_table_mapping[date_str].append(table)
                            
                        except Exception as e:
                            print(f"Warning: Could not parse date '{row[date_column]}' from {table}: {e}")
                
                print(f"✓ Processed table '{table}'")
                
            except Exception as e:
                print(f"✗ Error fetching dates from {table}: {e}")
                import traceback
                traceback.print_exc()
        
       
        dates_list = []
        for date_str in sorted(all_dates, reverse=True):
            dates_list.append({
                'date': date_str,
                'table_count': len(date_table_mapping[date_str]),
                'tables': date_table_mapping[date_str]
            })
        
        print(f"✓ Found {len(dates_list)} unique dates across all tables")
        
        return jsonify({
            'dates': dates_list,
            'total_count': len(dates_list),
            'success': True
        })
        
    except Exception as e:
        print(f"✗ Error fetching weekly data: {e}")
        import traceback
        traceback.print_exc()
        return jsonify({'error': str(e), 'success': False}), 500


@app.route('/api/weekly_data/delete', methods=['POST'])
def delete_weekly_data():
    """
    Delete data for a specific date from all 6 tables
    """
    try:
        data = request.get_json()
        date_to_delete = data.get('date')
        
        if not date_to_delete:
            return jsonify({'error': 'No date provided', 'success': False}), 400
        
        print(f"\n=== DELETING DATA FOR DATE: {date_to_delete} ===")
        
        tables_to_delete = [
            "daily_commando",
            "daily_st",
            "daily_percentage_docked",
            "commando_pivot_data",
            "st_pivot_data",
            "bay_alphabet_data"
        ]
        
        total_deleted = 0
        deletion_summary = {}
        
        for table in tables_to_delete:
            try:
                delete_res = supabase.table(table).delete().eq("date", date_to_delete).execute()
                
                deleted_count = len(delete_res.data) if delete_res.data else 0
                deletion_summary[table] = deleted_count
                total_deleted += deleted_count
                
                print(f"✓ Deleted {deleted_count} records from {table}")
                
            except Exception as e:
                print(f"✗ Error deleting from {table}: {e}")
                deletion_summary[table] = f"Error: {str(e)}"
        
        print(f"✓ Total records deleted: {total_deleted}")
        
        return jsonify({
            'success': True,
            'message': f'Deleted data for {date_to_delete} from all tables',
            'total_deleted': total_deleted,
            'deletion_summary': deletion_summary
        })
        
    except Exception as e:
        print(f"✗ Error deleting weekly data: {e}")
        import traceback
        traceback.print_exc()
        return jsonify({'error': str(e), 'success': False}), 500


@app.route('/api/commandos_weekly', methods=['GET'])
def get_commandos_weekly():
    try:
        res = supabase.table("commando_weekly").select("Dates, Count").order("id").execute()
        print(f"Commandos weekly data: {res.data}")
        return jsonify({'data': res.data, 'success': True})
    except Exception as e:
        print(f"Error fetching commandos weekly from DB: {e}")
        return jsonify({'data': [], 'error': str(e), 'success': False}), 500

@app.route('/api/st_weekly', methods=['GET'])
def get_st_weekly():
    try:
        res = supabase.table("st_weekly").select("Dates, Count").order("id").execute()
        print(f"ST weekly data: {res.data}")
        return jsonify({'data': res.data, 'success': True})
    except Exception as e:
        print(f"Error fetching ST weekly: {e}")
        return jsonify({'data': [], 'error': str(e), 'success': False}), 500

@app.route('/api/manage_weeks/delete', methods=['POST'])
def delete_weekly_data_old():
    try:
        request_data = request.get_json()
        date_range_to_delete = request_data.get('date_range')
        if not date_range_to_delete:
            return jsonify({'error': 'No date range provided', 'success': False}), 400

        tables = ["daily_commando", "daily_st", "daily_percentage_docked", "commando_pivot_data", "st_pivot_data"]
        for table in tables:
            supabase.table(table).delete().eq("flightDate", date_range_to_delete).execute()

        return jsonify({'success': True, 'message': f'Deleted date {date_range_to_delete} from Database'})
    except Exception as e:
        print(f"Error deleting weekly data: {e}")
        return jsonify({'error': str(e), 'success': False}), 500


@app.route('/api/processed_files', methods=['GET'])
def get_processed_files():
    try:
        res = supabase.storage.from_('processed').list()
        
        files = []
        for item in res:
            files.append({
                'name': item.get('name'),
                'size': item.get('metadata', {}).get('size', 0),
                'modified': item.get('updated_at')
            })
        
        files.sort(key=lambda x: x['modified'], reverse=True)
        
        print(f"Found {len(files)} files in Supabase bucket")
        return jsonify({'files': files, 'success': True})
    except Exception as e:
        print(f"Error fetching processed files from Supabase: {e}")
        import traceback
        traceback.print_exc()
        return jsonify({'error': str(e), 'success': False}), 500


@app.route('/api/processed_files/download/<filename>', methods=['GET'])
def download_processed_file(filename):
    """Download a file directly from Supabase bucket"""
    try:
        if '..' in filename or '/' in filename or '\\' in filename:
            return jsonify({'error': 'Invalid filename', 'success': False}), 400
        
        print(f"Downloading file from Supabase: {filename}")
        
        response = supabase.storage.from_('processed').download(filename)
        
        if not response:
            print(f"File not found in Supabase: {filename}")
            return jsonify({'error': 'File not found', 'success': False}), 404
        
        print(f"Successfully retrieved file from Supabase: {filename}")
        
        return send_file(
            io.BytesIO(response),
            as_attachment=True,
            download_name=filename,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
    except Exception as e:
        print(f"Error downloading file from Supabase: {e}")
        import traceback
        traceback.print_exc()
        return jsonify({'error': f'Error: {str(e)}', 'success': False}), 500


@app.route('/api/processed_files/delete', methods=['POST'])
def delete_processed_file():
    """Delete a file from Supabase bucket"""
    try:
        data = request.get_json()
        filename = data.get('filename')
        
        if not filename:
            return jsonify({'error': 'No filename provided', 'success': False}), 400
        
        if '..' in filename or '/' in filename or '\\' in filename:
            return jsonify({'error': 'Invalid filename', 'success': False}), 400
        
        print(f"Deleting file from Supabase: {filename}")
        supabase.storage.from_('processed').remove([filename])
        
        print(f"Successfully deleted: {filename}")
        return jsonify({'success': True, 'message': f'Deleted {filename}'})
    except Exception as e:
        print(f"Error deleting file from Supabase: {e}")
        import traceback
        traceback.print_exc()
        return jsonify({'error': str(e), 'success': False}), 500

def upload_to_supabase(local_path, filename):
    with open(local_path, 'rb') as f:
        res = supabase.storage.from_('processed').upload(
            path=filename,
            file=f,
            file_options={
                "content-type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", 
                "x-upsert": "true"  
            }
        )
    print(f"✓ Uploaded to Supabase: {filename} (overwrote existing if present)")
    return res

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)