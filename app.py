import os, sys, json, time, math, threading, shutil
from datetime import datetime, timedelta
from dotenv import load_dotenv

# --- WYKRYWANIE TRYBU CLI ---
CLI_MODE = False
if len(sys.argv) > 1 and '--send-mail' in sys.argv:
    CLI_MODE = True

# Importy wspólne dla obu trybów
import pandas as pd
import win32com.client
import pythoncom

# Streamlit tylko w trybie normalnym
if not CLI_MODE:
    import streamlit as st

# --- LOGGING ---
LOG_FILE = "server_log.txt"
MAX_LOG_SIZE = 5 * 1024 * 1024  # 5 MB
MAX_LOG_FILES = 5  # server_log.txt + 4 archiwów

def rotate_file_if_needed(filename: str, max_size: int, max_files: int):
    """Rotacja pliku jeśli przekroczył maksymalny rozmiar."""
    try:
        if os.path.exists(filename) and os.path.getsize(filename) > max_size:
            # Użyj tymczasowego pliku do logowania błędu rotacji
            temp_log = f"{filename}.rotation_error"
            try:
                with open(temp_log, 'a', encoding='utf-8') as f:
                    f.write(f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')} - INFO - Rozpoczynam rotację pliku {filename} (przekroczono {max_size//1024//1024}MB)\n")
                
                # Usuń najstarsze archiwum jeśli istnieje
                oldest_file = f"{filename}.{max_files-1}"
                if os.path.exists(oldest_file):
                    os.remove(oldest_file)
                    with open(temp_log, 'a', encoding='utf-8') as f:
                        f.write(f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')} - INFO - Usunięto najstarsze archiwum: {oldest_file}\n")
                
                # Przesuń istniejące archiwa
                for i in range(max_files - 2, 0, -1):
                    old_name = f"{filename}.{i}"
                    new_name = f"{filename}.{i+1}"
                    if os.path.exists(old_name):
                        shutil.move(old_name, new_name)
                
                # Przenieś aktualny plik do .1
                shutil.move(filename, f"{filename}.1")
                
                with open(temp_log, 'a', encoding='utf-8') as f:
                    f.write(f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')} - INFO - Rotacja pliku {filename} zakończona pomyślnie\n")
                
                # Przenieś logi rotacji do nowego pliku
                if os.path.exists(temp_log):
                    with open(temp_log, 'r', encoding='utf-8') as src:
                        content = src.read()
                    os.remove(temp_log)
                    with open(filename, 'w', encoding='utf-8') as dst:
                        dst.write(content)
                
            except Exception as e:
                # Jeśli coś pójdzie nie tak, zapisz błąd
                if os.path.exists(temp_log):
                    with open(temp_log, 'a', encoding='utf-8') as f:
                        f.write(f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')} - ERROR - Błąd rotacji: {e}\n")
                    shutil.move(temp_log, filename)
                else:
                    with open(filename, 'a', encoding='utf-8') as f:
                        f.write(f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')} - ERROR - Błąd rotacji: {e}\n")
    except Exception as e:
        # Ostateczny fallback - ignoruj błąd rotacji
        pass

def rotate_logs_if_needed():
    """Rotacja głównego pliku logów aplikacji."""
    rotate_file_if_needed(LOG_FILE, MAX_LOG_SIZE, MAX_LOG_FILES)

def log_event(level: str, message: str):
    """Zapisz zdarzenie do pliku logów."""
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    log_line = f"{timestamp} - {level} - {message}\n"
    
    # Najpierw próbuj zapisać do pliku
    try:
        with open(LOG_FILE, 'a', encoding='utf-8') as f:
            f.write(log_line)
        # Po zapisie sprawdź rotację
        try:
            rotate_logs_if_needed()
        except Exception as rotate_error:
            # Jeśli rotacja się nie powiedzie, tylko zapisz błąd
            with open(LOG_FILE, 'a', encoding='utf-8') as f:
                f.write(f"{timestamp} - WARNING - Błąd rotacji logów: {rotate_error}\n")
    except Exception as e:
        # Jeśli logowanie się nie powiedzie, wyświetl błąd w konsoli
        print(f"Nie udało się zapisać logu: {e}")
        # Spróbuj utworzyć plik od nowa jeśli nie istnieje
        try:
            with open(LOG_FILE, 'w', encoding='utf-8') as f:
                f.write(f"{timestamp} - INFO - Utworzono nowy plik logów\n")
                f.write(log_line)
        except:
            pass

load_dotenv()

# Rotacja logów przy starcie aplikacji
rotate_logs_if_needed()

# Usuń stary plik loga strażnika jeśli istnieje
if os.path.exists("guard_log.txt"):
    try:
        os.remove("guard_log.txt")
        log_event("INFO", "Usunięto stary plik loga strażnika: guard_log.txt")
    except Exception as e:
        log_event("WARNING", f"Nie udało się usunąć guard_log.txt: {e}")

URL = os.getenv("EXCEL_PATH")
SETS_FILE = "sets_cache.json"
DMA_CONFIG = "dma_config.json"
BUFFER_FILE = "today_buffer.json"
ACTIVE_FILE = "active_task.json"
META_CACHE_FILE = "meta_cache.json"
EMAIL_CACHE_FILE = "email_cache.json"
MANAGER_EMAIL = os.getenv("MANAGER_EMAIL")
USER_INITIALS = os.getenv("USER_INITIALS")
PRIVATE_EMAIL = os.getenv("PRIVATE_EMAIL", "")

# Lock dla operacji Excel (COM nie jest thread-safe)
excel_lock = threading.Lock()

# --- LOGIKA LOKALNA ---
def load_json(path, default=None):
    if default is None: default = []
    if os.path.exists(path):
        try:
            with open(path, "r", encoding="utf-8") as f: return json.load(f)
        except: return default
    return default

def save_json(path, data):
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=4)

def round_15(dt):
    # Zaokrąglenie do najbliższych 15 minut (granica 7.5) - użycie integera
    mins = (dt.minute + 7) // 15 * 15
    return dt + timedelta(minutes=mins - dt.minute) - timedelta(seconds=dt.second, microseconds=dt.microsecond)

def floor_15(dt):
    # Zaokrąglenie w dół do pełnych 15 minut
    mins = (dt.minute // 15) * 15
    return dt.replace(minute=mins, second=0, microsecond=0)

def get_rounded_hours(start_str, end_str):
    try:
        t_start = datetime.strptime(start_str, "%H:%M")
        t_end = datetime.strptime(end_str, "%H:%M")
        if t_end >= t_start:
            raw = (t_end - t_start).total_seconds() / 3600
            return max(0.25, math.ceil(raw * 4) / 4)
        return 0.25
    except: 
        return 0.25

# --- WALIDACJA POTENCJAŁU ---
def get_remaining_potential(pot_daily):
    buffer = load_json(BUFFER_FILE)
    used_h = sum(t['hours'] for t in buffer)
    return max(0.0, pot_daily - used_h)

# --- AUTO-CLEANUP STARYCH DANYCH ---
today_str = datetime.now().strftime("%Y-%m-%d")
raw_buffer = load_json(BUFFER_FILE)
valid_buffer = [t for t in raw_buffer if t.get('date') == today_str]
if len(valid_buffer) != len(raw_buffer):
    removed = len(raw_buffer) - len(valid_buffer)
    log_event("INFO", f"Usunięto {removed} starych wpisów z bufora")
    save_json(BUFFER_FILE, valid_buffer)

active_data = load_json(ACTIVE_FILE, None)
if active_data and active_data.get('date') != today_str:
    log_event("INFO", "Usunięto stare aktywne zadanie (z innego dnia)")
    if os.path.exists(ACTIVE_FILE): os.remove(ACTIVE_FILE)

# --- ZAMYKANIE AKTYWNEGO ZADANIA ---
def close_active_task(end_time_str):
    active = load_json(ACTIVE_FILE, None)
    if active:
        # Sprawdź, czy zadanie jest z dzisiaj
        if active.get('date') != today_str:
            # Usuń plik, nie dodawaj do bufora
            if os.path.exists(ACTIVE_FILE): os.remove(ACTIVE_FILE)
            return
        diff = get_rounded_hours(active['start'], end_time_str)
        buffer = load_json(BUFFER_FILE)
        buffer.append({
            "date": active['date'],
            "hours": diff,
            "proj": active['proj'], "char": active['char'], "opis": active['opis']
        })
        save_json(BUFFER_FILE, buffer)
        log_event("INFO", f"Zamknięto zadanie: {active['proj']} ({active['char']}) - {diff}h")
        if os.path.exists(ACTIVE_FILE): os.remove(ACTIVE_FILE)

# --- SILNIK EXCEL (ATOMOWY ZAPIS) ---
def scan_excel_for_sets(days_back=90):
    """
    Skanuje Excel w poszukiwaniu istniejących setów (F-H-I) z datami.
    
    Parameters:
    days_back (int): Liczba dni wstecz do skanowania (domyślnie 90)
    """
    log_event("INFO", f"Rozpoczecie skanowania Excela dla ostatnich {days_back} dni")
    
    # Użyj locka dla operacji Excel (COM nie jest thread-safe)
    with excel_lock:
        pythoncom.CoInitialize()
        excel = None
        wb = None
        retries = 3
        
        while retries > 0:
            try:
                excel = win32com.client.Dispatch("Excel.Application")
                excel.Visible = False
                excel.DisplayAlerts = False
                wb = excel.Workbooks.Open(URL)
                time.sleep(2)
                ws = wb.Worksheets("CAŁY ROK")
                
                # Znajdź ostatni wiersz w kolumnie A (daty)
                last_row = ws.Cells(ws.Rows.Count, 1).End(-4162).Row
                
                            # Oblicz datę graniczną (np. 90 dni wstecz od dzisiaj) - usuwamy strefę czasową
                cutoff_date = datetime.now().replace(tzinfo=None) - timedelta(days=days_back)
                
                # Zbierz dane z kolumn: A (data), F (projekt), H (charakter), I (opis)
                sets_with_dates = []
                for row in range(2, last_row + 1):  # Pomijamy nagłówek (zakładając, że wiersz 1 to nagłówki)
                    # Pobierz datę
                    date_val = ws.Cells(row, 1).Value
                    # Pobierz wartości F, H, I
                    f_val = ws.Cells(row, 6).Value  # kolumna F
                    h_val = ws.Cells(row, 8).Value  # kolumna H
                    i_val = ws.Cells(row, 9).Value  # kolumna I
                    
                    # Jeśli wszystkie trzy pola są wypełnione, dodaj do listy
                    if f_val and h_val and i_val:
                        # Konwersja na stringi i czyszczenie
                        f_str = str(f_val).strip()
                        h_str = str(h_val).strip()
                        i_str = str(i_val).strip()
                        
                                            # Konwersja daty - może być datetime lub string
                        date_obj = None
                        if hasattr(date_val, 'year'):  # obiekt datetime
                            date_obj = date_val
                            # Upewnij się, że data jest bez strefy czasowej (offset-naive)
                            if date_obj.tzinfo is not None:
                                date_obj = date_obj.replace(tzinfo=None)
                        else:
                            # Spróbuj sparsować string do daty
                            try:
                                if isinstance(date_val, str):
                                    date_obj = datetime.strptime(date_val, "%Y-%m-%d")
                                else:
                                    # Jeśli nie string, użyj dzisiejszej daty jako fallback
                                    date_obj = datetime.now().replace(tzinfo=None)
                            except:
                                date_obj = datetime.now().replace(tzinfo=None)
                        
                        # Pomijaj wiersze starsze niż cutoff_date
                        if date_obj < cutoff_date:
                            continue
                        
                        sets_with_dates.append({
                            "date": date_obj,
                            "F": f_str,
                            "H": h_str,
                            "I": i_str
                        })
            
                                    
                # Agregacja: dla każdej unikalnej kombinacji F-H-I znajdź najświeższą datę
                unique_sets = {}
                for item in sets_with_dates:
                    key = (item["F"], item["H"], item["I"])
                    if key not in unique_sets:
                        unique_sets[key] = {"F": item["F"], "H": item["H"], "I": item["I"], "latest_date": item["date"]}
                    else:
                        # Aktualizuj datę, jeśli znaleziono nowszą
                        if item["date"] > unique_sets[key]["latest_date"]:
                            unique_sets[key]["latest_date"] = item["date"]
                
                # Sortowanie według daty (najnowsze pierwsze)
                # Użyjemy tymczasowej listy z datami do sortowania
                temp_for_sorting = []
                for key, value in unique_sets.items():
                    temp_for_sorting.append({
                        "F": value["F"],
                        "H": value["H"],
                        "I": value["I"],
                        "latest_date": value["latest_date"]
                    })
                
                temp_for_sorting.sort(key=lambda x: x["latest_date"], reverse=True)
                
                # Przygotuj wynikową listę z datami (zachowujemy datę jako string)
                result_list = []
                for item in temp_for_sorting:
                    # Konwertuj datę na string w formacie YYYY-MM-DD
                    date_str = item["latest_date"].strftime("%Y-%m-%d")
                    result_list.append({
                        "F": item["F"],
                        "H": item["H"],
                        "I": item["I"],
                        "date": date_str  # Dodajemy datę
                    })
                
                                        # Wyczyść istniejący cache i zapisz nowe sety
                save_json(SETS_FILE, result_list)
                log_event("INFO", f"Znaleziono {len(result_list)} unikalnych setów z ostatnich {days_back} dni")
                
                return True, f"Znaleziono {len(result_list)} unikalnych setów z ostatnich {days_back} dni"
                
            except Exception as e:
                retries -= 1
                time.sleep(2)
                if retries == 0:
                    log_event("ERROR", f"Błąd podczas skanowania Excela: {e}")
                    return False, f"Błąd podczas skanowania Excela: {e}"
                # Kontynuuj próbę
            finally:
                try:
                    if wb:
                        try:
                            wb.Close(False)
                        except Exception as close_error:
                            log_event("WARNING", f"Błąd przy zamykaniu skoroszytu Excel w scan_excel_for_sets: {close_error}")
                        wb = None
                    if excel:
                        try:
                            excel.Quit()
                        except Exception as quit_error:
                            log_event("WARNING", f"Błąd przy zamykaniu Excel w scan_excel_for_sets: {quit_error}")
                        excel = None
                except Exception as final_error:
                    log_event("WARNING", f"Błąd w sekcji finally scan_excel_for_sets: {final_error}")
        
        try:
            pythoncom.CoUninitialize()
        except Exception as com_error:
            log_event("WARNING", f"Błąd przy zwalnianiu COM w scan_excel_for_sets: {com_error}")
        return False, "Nie udało się skanować Excela po 3 próbach"


def excel_worker(push_data=None, get_meta=False, w_s="09:00", w_e="17:00"):
    log_event("INFO", f"Excel worker: push_data={push_data is not None}, get_meta={get_meta}, w_s={w_s}, w_e={w_e}")
    
    # Użyj locka dla operacji Excel (COM nie jest thread-safe)
    with excel_lock:
        pythoncom.CoInitialize()
        excel = None
        wb = None
        retries = 3
        res = None
        
        for attempt in range(retries):
            try:
                excel = win32com.client.Dispatch("Excel.Application")
                excel.Visible = False
                excel.DisplayAlerts = False
                wb = excel.Workbooks.Open(URL)
                time.sleep(2)
                ws = wb.Worksheets("CAŁY ROK")
                today = datetime.now()
                
                last_row_a = ws.Cells(ws.Rows.Count, 1).End(-4162).Row
                dates_col = ws.Range(ws.Cells(1,1), ws.Cells(last_row_a, 1)).Value
                today_rows = [i + 1 for i, row in enumerate(dates_col) if row[0] and hasattr(row[0], 'year') and 
                              row[0].year == today.year and row[0].month == today.month and row[0].day == today.day]

                if get_meta:
                    def get_l(name):
                        try:
                            return [str(item[0]).strip() for item in excel.Range(name).Value if item[0]]
                        except:
                            return []
                    res = {"P": get_l("Projekty[Lista_projektów]"), "C": get_l("CHARAKTER"), "Z": get_l("ZAKRES")}
                    break

                if push_data is not None:
                    # Jeśli brak wiersza z dzisiejszą datą, dodajemy nowy wiersz na końcu
                    if not today_rows:
                        new_row = last_row_a + 1
                        ws.Cells(new_row, 1).Value = today
                        today_rows = [new_row]
                        # Resetujemy kolumny dla tego wiersza
                        for col in [3, 4, 6, 8, 9, 10]:
                            ws.Cells(new_row, col).Value = None
                    else:
                        # KROK 1: CAŁKOWITY RESET DZISIEJSZEGO OBSZARU
                        for r in today_rows:
                            for col in [3, 4, 6, 8, 9, 10]:  # C, D, F, H, I, J
                                ws.Cells(r, col).Value = None

                    # KROK 2: WPISANIE RAM DNIA (tylko w pierwszym wierszu)
                    ws.Cells(today_rows[0], 3).Value = w_s
                    ws.Cells(today_rows[0], 4).Value = w_e
                    
                    # KROK 3: AGREGACJA ZADAŃ (Lustrzana)
                    agg_tasks = {}
                    for t in push_data:
                        key = (str(t['proj']).strip(), str(t['char']).strip(), str(t['opis']).strip())
                        if key not in agg_tasks:
                            agg_tasks[key] = t.copy()
                        else:
                            agg_tasks[key]['hours'] += t['hours']
                    
                    # KROK 4: ZAPIS OD NOWA
                    # Jeśli za mało wierszy, dodaj nowe wiersze po ostatnim dzisiejszym wierszu
                    if len(agg_tasks) > len(today_rows):
                        last_row = today_rows[-1]
                        rows_needed = len(agg_tasks) - len(today_rows)
                        for i in range(rows_needed):
                            new_row = last_row + i + 1
                            # Wstawienie pustego wiersza (kopiowanie formatu?)
                            ws.Rows(new_row).Insert()
                            # Skopiowanie daty z poprzedniego wiersza (kolumna A)
                            ws.Cells(new_row, 1).Value = ws.Cells(last_row, 1).Value
                            today_rows.append(new_row)
                    
                    for idx, task in enumerate(agg_tasks.values()):
                        r = today_rows[idx]
                        ws.Cells(r, 6).Value = task['proj']
                        ws.Cells(r, 8).Value = task['char']
                        ws.Cells(r, 9).Value = task['opis']
                        ws.Cells(r, 10).Value = task['hours']

                    wb.Save()
                    log_event("INFO", "Zapis do Excela zakończony sukcesem")
                    
                    # Zamiast natychmiastowej aktualizacji cache, oznacz jako potrzebną aktualizację
                    # Cache zostanie zaktualizowany przy następnym wysłaniu maila
                    log_event("INFO", "Dane zapisane do Excela, cache wymaga aktualizacji")
                    res = True
                    break
                    
            except Exception as e:
                log_event("WARNING", f"Próba {attempt + 1}/{retries} nieudana: {e}")
                if attempt == retries - 1:  # Ostatnia próba
                    log_event("ERROR", f"Błąd Excela po {retries} próbach: {e}")
                    if not CLI_MODE:
                        st.error(f"⚠️ Błąd Excela: {e}")
                else:
                    time.sleep(2)  # Czekaj przed kolejną próbą
            finally:
                # Zawsze spróbuj zamknąć Excel
                try:
                    if wb:
                        try:
                            wb.Close(False)
                        except Exception as close_error:
                            log_event("WARNING", f"Błąd przy zamykaniu skoroszytu Excel (może być już zamknięty): {close_error}")
                    if excel:
                        try:
                            excel.Quit()
                        except Exception as quit_error:
                            log_event("WARNING", f"Błąd przy zamykaniu Excel (może być już zamknięty): {quit_error}")
                except Exception as final_error:
                    log_event("WARNING", f"Błąd w sekcji finally excel_worker: {final_error}")
        
        # Zwolnij COM po zakończeniu wszystkich prób
        try:
            pythoncom.CoUninitialize()
        except Exception as com_error:
            log_event("WARNING", f"Błąd przy zwalnianiu COM: {com_error}")
        
        return res

# --- CACHE DLA RAPORTU E-MAIL ---
def update_email_cache():
    """Aktualizuj cache danych do raportu email."""
    log_event("INFO", "Aktualizacja cache dla raportu email")
    
    # Użyj locka dla operacji Excel (COM nie jest thread-safe)
    with excel_lock:
        pythoncom.CoInitialize()
        excel = None
        wb = None
        data = None
        
        try:
            # Pobierz dane z Excela
            excel = win32com.client.Dispatch("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False
            wb = excel.Workbooks.Open(URL)
            time.sleep(1)
            ws = wb.Worksheets("CAŁY ROK")
            
            last_r = ws.Cells(ws.Rows.Count, 1).End(-4162).Row
            data = ws.Range(ws.Cells(1, 1), ws.Cells(last_r, 10)).Value
            
            # Przetwarzanie danych
            raw_rows = list(data[1:])
            cleaned_rows = []
            
            for row in raw_rows:
                row_list = list(row)
                val = row_list[0]
                if val and hasattr(val, 'year') and hasattr(val, 'month') and hasattr(val, 'day'):
                    row_list[0] = f"{val.year}-{val.month:02d}-{val.day:02d}"
                cleaned_rows.append(row_list)

            df = pd.DataFrame(cleaned_rows, columns=data[0])
            df.iloc[:, 0] = pd.to_datetime(df.iloc[:, 0], errors='coerce')
            df = df.dropna(subset=[df.columns[0]]) 
            
            # Filtrowanie pustych dni i weekendów bez pracy
            df = df.dropna(subset=[df.columns[5]])
            df = df[df.iloc[:, 5].astype(str).str.strip() != ""]
            df = df[df.iloc[:, 5].astype(str).str.strip() != "None"]
            
            # Sortowanie
            df = df.sort_values(by=df.columns[0], ascending=False)
            
            # Ostatnie 20 dni
            top_20_dates = df[df.columns[0]].unique()[:20]
            df = df[df[df.columns[0]].isin(top_20_dates)]
            
            # Przygotowanie danych do cache
            cache_data = {
                "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "data": []
            }
            
            for date, group in df.groupby(df.columns[0], sort=False):
                for _, row in group.iterrows():
                    cache_data["data"].append({
                        "date": date.strftime("%Y-%m-%d"),
                        "project": str(row.iloc[5]) if pd.notna(row.iloc[5]) else "",
                        "character": str(row.iloc[7]) if pd.notna(row.iloc[7]) else "",
                        "description": str(row.iloc[8]) if pd.notna(row.iloc[8]) else "",
                        "hours": float(row.iloc[9]) if pd.notna(row.iloc[9]) else 0.0
                    })
            
            # Zapisz cache
            save_json(EMAIL_CACHE_FILE, cache_data)
            log_event("INFO", f"Zaktualizowano cache dla raportu email: {len(cache_data['data'])} wpisów")
            
        except Exception as e:
            log_event("ERROR", f"Błąd aktualizacji cache email: {e}")
            import traceback
            error_details = traceback.format_exc()
            log_event("ERROR", f"Szczegóły błędu cache: {error_details}")
        finally:
            try:
                if wb:
                    try: 
                        wb.Close(False)
                    except Exception as close_error:
                        log_event("WARNING", f"Błąd przy zamykaniu skoroszytu Excel w update_email_cache: {close_error}")
                if excel:
                    try:
                        excel.Quit()
                    except Exception as quit_error:
                        log_event("WARNING", f"Błąd przy zamykaniu Excel w update_email_cache: {quit_error}")
            except Exception as final_error:
                log_event("WARNING", f"Błąd w sekcji finally update_email_cache: {final_error}")
            finally:
                try:
                    pythoncom.CoUninitialize()
                except Exception as com_error:
                    log_event("WARNING", f"Błąd przy zwalnianiu COM w update_email_cache: {com_error}")


# --- WYSYŁKA RAPORTU E-MAIL ---
def send_formatted_mail():
    """Wyślij raport e-mail z ostatnimi dniami pracy (używa cache)."""
    log_event("INFO", "Rozpoczęcie wysyłania raportu e-mail z cache")
    log_event("INFO", f"EXCEL_PATH: {URL}")
    log_event("INFO", f"MANAGER_EMAIL: {MANAGER_EMAIL}")
    log_event("INFO", f"USER_INITIALS: {USER_INITIALS}")
    
    # --- CZĘŚĆ 1: POBRANIE DANYCH Z CACHE ---
    cache_data = load_json(EMAIL_CACHE_FILE, None)
    
    # Sprawdź, czy cache istnieje i jest świeży (mniej niż 1 godzina)
    cache_valid = False
    if cache_data and "timestamp" in cache_data and "data" in cache_data:
        try:
            cache_time = datetime.strptime(cache_data["timestamp"], "%Y-%m-%d %H:%M:%S")
            time_diff = datetime.now() - cache_time
            if time_diff.total_seconds() < 3600:  # 1 godzina
                cache_valid = True
                log_event("INFO", f"Używam cache z {cache_data['timestamp']} ({len(cache_data['data'])} wpisów)")
            else:
                log_event("INFO", f"Cache jest za stary: {time_diff.total_seconds()//60} minut")
        except Exception as e:
            log_event("WARNING", f"Błąd parsowania czasu cache: {e}")
    
    # Jeśli cache nie jest ważny, zaktualizuj go
    if not cache_valid:
        log_event("INFO", "Aktualizuję cache danych z Excela")
        update_email_cache()
        cache_data = load_json(EMAIL_CACHE_FILE, None)
        
        if not cache_data or "data" not in cache_data:
            log_event("ERROR", "Nie udało się załadować danych z cache nawet po aktualizacji")
            return
    
    # --- CZĘŚĆ 2: PRZYGOTOWANIE HTML Z CACHE ---
    if not cache_data["data"]:
        log_event("WARNING", "Cache jest pusty, nie ma danych do wysłania")
        return
    
    # Grupowanie danych według daty
    from collections import defaultdict
    grouped_by_date = defaultdict(list)
    for item in cache_data["data"]:
        grouped_by_date[item["date"]].append(item)
    
    # Sortowanie dat (najnowsze pierwsze) - ogranicz do 20 dni
    sorted_dates = sorted(grouped_by_date.keys(), reverse=True)[:20]
    
    days_pl = ["poniedziałek", "wtorek", "środa", "czwartek", "piątek", "sobota", "niedziela"]
    html = "<html><body style='font-family: Calibri, sans-serif;'>"
    
    for date_str in sorted_dates:
        date_obj = datetime.strptime(date_str, "%Y-%m-%d")
        day_str = days_pl[date_obj.weekday()]
        html += f"<p><b>{date_str}; {day_str}</b></p>"
        
        # Tabela dla danej daty
        html += "<table border='1' cellpadding='5' cellspacing='0' style='border-collapse: collapse;'>"
        html += "<tr><th>Projekt</th><th>Charakter</th><th>Opis</th><th>Godziny</th></tr>"
        
        for item in grouped_by_date[date_str]:
            html += f"<tr><td>{item['project']}</td><td>{item['character']}</td><td>{item['description']}</td><td>{item['hours']}h</td></tr>"
        
        html += "</table><br><br>"
    
    html += "</body></html>"

    # --- CZĘŚĆ 2: WYSYŁKA E-MAIL ---
    outlook = None
    try:
        # Użyj EnsureDispatch dla lepszej kompatybilności z Outlook
        log_event("INFO", "Próba połączenia z Outlook...")
        try:
            outlook = win32com.client.GetActiveObject("Outlook.Application")
            log_event("INFO", "Połączono z aktywnym Outlook")
        except:
            outlook = win32com.client.gencache.EnsureDispatch("Outlook.Application")
            log_event("INFO", "Uruchomiono nową instancję Outlook")
        
        # Daj Outlook czas na inicjalizację
        time.sleep(2)
        
        mail = outlook.CreateItem(0)  # 0 = olMailItem
        mail.Subject = f"#RAPORT - {USER_INITIALS} - {datetime.now().year}"
        mail.To = MANAGER_EMAIL
        
        # --- DODANE UDW Z PLIKU .ENV ---
        if PRIVATE_EMAIL:
            mail.BCC = PRIVATE_EMAIL
            
        mail.HTMLBody = html
        log_event("INFO", "Wysyłam email...")
        mail.Send()
        log_event("INFO", f"Raport wysłany pomyślnie na adres: {MANAGER_EMAIL}")
        if PRIVATE_EMAIL:
            log_event("INFO", f"Kopia UDW wysłana na adres: {PRIVATE_EMAIL}")
        print(f"Raport wysłany pomyślnie na adres: {MANAGER_EMAIL}!")
        if PRIVATE_EMAIL:
            print(f"Kopia UDW wysłana na adres: {PRIVATE_EMAIL}!")
    except Exception as e:
        log_event("ERROR", f"Błąd wysyłania e-maila: {e}")
        print(f"Błąd wysyłania e-maila: {e}")
        # Dodaj szczegółowe informacje o błędzie
        import traceback
        error_details = traceback.format_exc()
        log_event("ERROR", f"Szczegóły błędu: {error_details}")
    finally:
        # Zawsze zwolnij COM
        try:
            pythoncom.CoUninitialize()
            log_event("INFO", "COM zwolniony")
        except Exception as com_error:
            log_event("WARNING", f"Błąd podczas zwalniania COM: {com_error}")


def schedule_email(target_str: str):
    """Zaplanuj wysłanie e-maila na określoną godzinę."""
    try:
        target_time = datetime.strptime(target_str, "%H:%M").time()
        target_dt = datetime.combine(datetime.now().date(), target_time)
        now = datetime.now()
        
        if target_dt < now:
            # Jeśli godzina już minęła, wyślij natychmiast
            log_event("INFO", f"Godzina {target_str} już minęła, wysyłam natychmiast")
            try:
                send_formatted_mail()
            except Exception as e:
                log_event("ERROR", f"Błąd podczas natychmiastowego wysyłania email: {e}")
                st.error(f"Błąd wysyłania raportu: {e}")
            return
        
        # Oblicz opóźnienie w sekundach
        delay_seconds = (target_dt - now).total_seconds()
        
        log_event("INFO", f"Zaplanowano wysyłkę raportu na {target_str} (za {delay_seconds:.0f} sekund)")
        
        # Uruchom wątek z opóźnieniem
        timer = threading.Timer(delay_seconds, send_formatted_mail)
        timer.daemon = True  # Wątek daemon zakończy się z głównym programem
        timer.start()
        
        st.success(f"Raport zaplanowany na {target_str}! (za {delay_seconds//60:.0f} minut {delay_seconds%60:.0f} sekund)")
        
    except Exception as e:
        log_event("ERROR", f"Błąd planowania e-maila: {e}")
        st.error(f"Błąd planowania: {e}")


# --- OBSŁUGA TRYBU CLI ---
if CLI_MODE:
    print("TrackMyDay - wysyłka raportu mailowego")
    print(f"EXCEL_PATH: {URL}")
    print(f"MANAGER_EMAIL: {MANAGER_EMAIL}")
    print(f"USER_INITIALS: {USER_INITIALS}")
    if PRIVATE_EMAIL:
        print(f"PRIVATE_EMAIL: {PRIVATE_EMAIL}")
    print()
    send_formatted_mail()
    print("Raport wysłany. Sprawdź Outlook.")
    sys.exit(0)

# --- INTERFEJS ---
log_event("INFO", "Uruchomienie aplikacji TrackMyDay")
st.set_page_config(page_title="TrackMyDay", layout="wide")
st.title("⚡ EWIDENCJA CZASU PRACY")

now_r = round_15(datetime.now()).strftime("%H:%M")

if 'manual_start' not in st.session_state: st.session_state.manual_start = floor_15(datetime.now()).strftime("%H:%M")
if 'manual_end' not in st.session_state: st.session_state.manual_end = now_r
if 'last_set' not in st.session_state: st.session_state.last_set = None

# Ładowanie słowników z cache (nie automatycznie z Excela)
if 'meta' not in st.session_state:
    meta_cache = load_json(META_CACHE_FILE, {"P": [], "C": [], "Z": []})
    st.session_state.meta = meta_cache
    log_event("INFO", f"Załadowano słowniki z cache: {len(meta_cache['P'])} projektów, {len(meta_cache['C'])} charakterów, {len(meta_cache['Z'])} opisów")

with st.sidebar:
    st.header("⚙️ RAMY CZASOWE")
    w_s = st.text_input("Start dnia:", value="09:00")
    w_e = st.text_input("Koniec dnia", value="17:00")
    try:
        delta = datetime.strptime(w_e, "%H:%M") - datetime.strptime(w_s, "%H:%M")
        pot_daily = delta.total_seconds() / 3600
        if pot_daily <= 0:
            pot_daily = 8.0  # default if end <= start
    except: pot_daily = 8.0
    
    rem_pot = get_remaining_potential(pot_daily)
    
    # Uwzględnij aktywne zadanie (niezależnie czy DMA czy nie)
    active_task = load_json(ACTIVE_FILE, None)
    active_hours = 0
    if active_task:
        active_hours = get_rounded_hours(active_task['start'], now_r)
        rem_pot = max(0.0, rem_pot - active_hours)
    
    # --- PASEK POSTĘPU ---
    used_h = pot_daily - rem_pot
    prog = max(0.0, min(1.0, used_h / pot_daily))
    color = "green" if rem_pot > 0.5 else "orange" if rem_pot > 0 else "red"
    
    st.markdown(f"Pozostało: <b style='color:{color}'>{round(rem_pot, 2)} h</b>", unsafe_allow_html=True)
    st.progress(prog)
    st.caption(f"Wykorzystano {round(used_h, 2)}h z {pot_daily}h")
    
        # Zapisz wartości do session_state dla głównego kodu
    st.session_state.pot_daily = pot_daily
    st.session_state.rem_pot = rem_pot
    st.session_state.w_s = w_s
    st.session_state.w_e = w_e
    
    st.divider()
    if st.button("🔍 AKTUALIZUJ SETY Z EXCELA", use_container_width=True):
        with st.spinner("Skanowanie Excela w poszukiwaniu istniejących setów..."):
            success, message = scan_excel_for_sets()
            if success:
                log_event("INFO", f"Skanowanie Excela zakończone sukcesem: {message}")
                st.success(message)
                # Odśwież listę setów w session_state
                if 'sets_df' in st.session_state:
                    del st.session_state.sets_df
                st.rerun()
            else:
                log_event("ERROR", f"Skanowanie Excela nie powiodło się: {message}")
                st.error(message)
    
    if st.button("🔄 ODŚWIEŻ LISTY ROZWIJANE", use_container_width=True):
        with st.spinner("Pobieram aktualne listy z Excela..."):
            res = excel_worker(get_meta=True)
            if res:
                save_json(META_CACHE_FILE, res)
                st.session_state.meta = res
                log_event("INFO", f"Zaktualizowano słowniki: {len(res['P'])} projektów, {len(res['C'])} charakterów, {len(res['Z'])} opisów")
                st.success(f"Zaktualizowano listy! Projekty: {len(res['P'])}, Charaktery: {len(res['C'])}, Opisy: {len(res['Z'])}")
            else:
                log_event("ERROR", "Nie udało się pobrać słowników z Excela")
                st.error("Nie udało się pobrać słowników z Excela. Sprawdź połączenie z plikiem Excel.")

# Pobierz wartości obliczone w sidebar
try:
    pot_daily = st.session_state.pot_daily
    rem_pot = st.session_state.rem_pot
    w_s = st.session_state.w_s
    w_e = st.session_state.w_e
except:
    pot_daily = 8.0
    rem_pot = pot_daily
    w_s = "09:00"
    w_e = "17:00"

# --- WYBÓR SETU ---
df_sets = pd.DataFrame(load_json(SETS_FILE, []))
# Jeśli DataFrame nie jest pusty, posortuj według daty (najnowsze pierwsze)
if not df_sets.empty:
    # Sprawdź, czy kolumna 'date' istnieje
    if 'date' in df_sets.columns:
        # Konwertuj daty na datetime dla poprawnego sortowania
        try:
            df_sets['date'] = pd.to_datetime(df_sets['date'])
        except:
            # Jeśli konwersja się nie uda, dodaj kolumnę z dzisiejszą datą
            df_sets['date'] = pd.Timestamp(today_str)
        df_sets = df_sets.sort_values('date', ascending=False)
    else:
        # Jeśli brak kolumny date, dodaj ją z dzisiejszą datą
        df_sets['date'] = pd.Timestamp(today_str)
    
    # Utwórz listę opcji
    options = ["--- Nowy SET ---"] + (df_sets['F'] + " | " + df_sets['H'] + " | " + df_sets['I']).tolist()
else:
    options = ["--- Nowy SET ---"]

sel_set = st.selectbox("Wybierz zadanie z bazy:", options)

if sel_set != st.session_state.last_set:
    st.session_state.last_set = sel_set
    st.session_state.manual_start = floor_15(datetime.now()).strftime("%H:%M")
    st.session_state.manual_end = now_r

p_idx, c_idx, o_idx = 0, 0, 0
if sel_set != "--- Nowy SET ---":
    row = df_sets.iloc[options.index(sel_set) - 1]
    m = st.session_state.meta
    if str(row['F']) in m['P']: p_idx = m['P'].index(str(row['F']))
    if str(row['H']) in m['C']: c_idx = m['C'].index(str(row['H']))
    if str(row['I']) in m['Z']: o_idx = m['Z'].index(str(row['I']))

col_f1, col_f2, col_f3 = st.columns(3)
p_sel = col_f1.selectbox("Projekt (F)", st.session_state.meta['P'], index=p_idx)
c_sel = col_f2.selectbox("Charakter (H)", st.session_state.meta['C'], index=c_idx)
o_sel = col_f3.selectbox("Opis (I)", st.session_state.meta['Z'], index=o_idx)

st.divider()

# --- AKTYWNE ZADANIE ---
active_task = load_json(ACTIVE_FILE, None)
if active_task:
    st.success(f"🟢 **W TRAKCIE:** {active_task['proj']} (Start: {active_task['start']})")
    if st.button("🛑 ZAKOŃCZ OBECNE ZADANIE", type="primary", use_container_width=True):
        close_active_task(now_r)
        st.rerun()
    st.divider()

st.write("⏱️ **Zarządzanie Czasem**")
t_c1, t_c2 = st.columns(2)
f_start = t_c1.text_input("Ręczny Start:", value=st.session_state.manual_start)
f_end = t_c2.text_input("Ręczny Koniec:", value=st.session_state.manual_end)

b1, b2, b3 = st.columns(3)

if b1.button("▶️ START ZADANIA", use_container_width=True):
    # Walidacja formatu czasu
    try:
        datetime.strptime(f_start, "%H:%M")
        start_valid = True
    except ValueError:
        st.error("Nieprawidłowy format czasu startu! Użyj HH:MM")
        start_valid = False
    
    if start_valid and rem_pot < 0.25:
        st.error("Brak wolnego czasu (minimum 0.25h)!")
    elif start_valid:
        close_active_task(now_r)
        save_json(ACTIVE_FILE, {"date": today_str, "proj": p_sel, "char": c_sel, "opis": o_sel, "start": f_start})
        log_event("INFO", f"Rozpoczęto zadanie: {p_sel} ({c_sel}) - start {f_start}")
        cache = load_json(SETS_FILE)
        # Sprawdź, czy set już istnieje (porównanie bez daty)
        exists = False
        for item in cache:
            if item.get("F") == p_sel and item.get("H") == c_sel and item.get("I") == o_sel:
                exists = True
                # Aktualizuj datę na dzisiejszą
                item["date"] = today_str
                break
        
        if not exists:
            new_s = {"F": p_sel, "H": c_sel, "I": o_sel, "date": today_str}
            cache.append(new_s)
        
        save_json(SETS_FILE, cache)
        st.rerun()

if b2.button("💾 ZAPISZ RĘCZNIE", use_container_width=True):
    # Walidacja formatów czasów
    try:
        datetime.strptime(f_start, "%H:%M")
        datetime.strptime(f_end, "%H:%M")
        times_valid = True
    except ValueError:
        st.error("Nieprawidłowy format czasu! Użyj HH:MM")
        times_valid = False
    
    if times_valid:
        needed = get_rounded_hours(f_start, f_end)
        if needed > rem_pot:
            st.error(f"Brakuje {needed - rem_pot}h!")
        else:
            buffer = load_json(BUFFER_FILE)
            buffer.append({"date": today_str, "hours": needed, "proj": p_sel, "char": c_sel, "opis": o_sel})
            save_json(BUFFER_FILE, buffer)
            log_event("INFO", f"Zapisano ręcznie: {p_sel} ({c_sel}) - {needed}h ({f_start}-{f_end})")
            st.session_state.manual_start = floor_15(datetime.now()).strftime("%H:%M"); st.session_state.manual_end = now_r
            st.rerun()

if b3.button("⭐ USTAW JAKO DMA", use_container_width=True):
    save_json(DMA_CONFIG, {"P": p_sel, "C": c_sel, "O": o_sel})
    log_event("INFO", f"Ustawiono DMA: {p_sel} ({c_sel}) - {o_sel}")
    st.success("DMA ustawione!"); st.rerun()

# --- PODGLĄD DNIA ---
# Pobierz wartości z sidebar (lub domyślne)
pot_daily = st.session_state.get('pot_daily', 8.0)
rem_pot = st.session_state.get('rem_pot', pot_daily)
w_s = st.session_state.get('w_s', '09:00')
w_e = st.session_state.get('w_e', '17:00')

st.subheader("📅 Podgląd Dnia")
buffer = load_json(BUFFER_FILE); dma_cfg = load_json(DMA_CONFIG, None)
active_hours = get_rounded_hours(active_task['start'], now_r) if active_task else 0

if buffer or dma_cfg or active_task:
    # Oblicz czas dla DMA (pozostały czas po wszystkich zadaniach i aktywnym)
    dma_h = max(0.0, rem_pot) if dma_cfg else 0

    h1, h2, h3, h4, h5, h6 = st.columns([1.5, 3, 2, 4, 1.5, 1])
    h1.markdown("**Typ**"); h2.markdown("**Projekt**"); h3.markdown("**Charakter**"); h4.markdown("**Opis**"); h5.markdown("**Godziny**"); h6.markdown("**Akcja**")
    st.markdown("<hr style='margin: 0px;'/>", unsafe_allow_html=True)
    
    if active_task:
        c1, c2, c3, c4, c5, c6 = st.columns([1.5, 3, 2, 4, 1.5, 1])
        c1.write("🟢 W Trakcie"); c2.write(active_task['proj']); c3.write(active_task['char']); c4.write(active_task['opis']); c5.write(f"~{active_hours}h"); c6.write("")
        st.markdown("<hr style='margin: 0px;'/>", unsafe_allow_html=True)

    for i, t in enumerate(buffer):
        c1, c2, c3, c4, c5, c6 = st.columns([1.5, 3, 2, 4, 1.5, 1])
        c1.write("Zadanie"); c2.write(t['proj']); c3.write(t['char']); c4.write(t['opis']); c5.write(f"{t['hours']}h")
        if c6.button("❌", key=f"del_{i}"):
            removed_task = buffer.pop(i)
            log_event("INFO", f"Usunięto zadanie z bufora: {removed_task['proj']} ({removed_task['char']}) - {removed_task['hours']}h")
            save_json(BUFFER_FILE, buffer)
            st.rerun()
        st.markdown("<hr style='margin: 0px;'/>", unsafe_allow_html=True)

    if dma_cfg and dma_h > 0:
        c1, c2, c3, c4, c5, c6 = st.columns([1.5, 3, 2, 4, 1.5, 1])
        c1.write("⭐ DMA"); c2.write(dma_cfg['P']); c3.write(dma_cfg['C']); c4.write(dma_cfg['O']); c5.write(f"{round(dma_h, 2)}h"); c6.write("")

st.divider()

# --- ZAKOŃCZ DZIEŃ ---
if st.button("🏁 ZAKOŃCZ DZIEŃ", use_container_width=True, type="primary"):
    with st.spinner("Synchronizacja z Excel..."):
        if active_task: close_active_task(w_e)
        buffer = load_json(BUFFER_FILE, [])
        if dma_cfg:
            # Usuń istniejące zadania DMA z bufora (zastąpimy je nowym z pełnym pozostałym czasem)
            buffer_without_dma = [t for t in buffer if not (t['proj'] == dma_cfg['P'] and t['char'] == dma_cfg['C'] and t['opis'] == dma_cfg['O'])]
            total_used = sum(t['hours'] for t in buffer_without_dma)
            final_rem = max(0.0, pot_daily - total_used)
            push_list = buffer_without_dma.copy()
            if final_rem > 0:
                push_list.append({"date": today_str, "hours": final_rem, "proj": dma_cfg['P'], "char": dma_cfg['C'], "opis": dma_cfg['O']})
        else:
            push_list = buffer.copy()
        
        if excel_worker(push_data=push_list, w_s=w_s, w_e=w_e):
            log_event("INFO", f"Zakończono dzień, zapisano {len(push_list)} zadań do Excela")
            
            target_str = (datetime.strptime(w_e, "%H:%M") - timedelta(minutes=5)).strftime("%H:%M")
            schedule_email(target_str)

            st.success(f"Zakończono dzień! Raport zaplanowany na {target_str}")