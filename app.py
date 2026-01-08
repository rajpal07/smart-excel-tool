import streamlit as st
import pandas as pd
import sys
import os
import tempfile

# Add current dir to path for imports
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

# Changed import from src to lib to avoid collision with root src folder
from lib.excel_utils import ExcelHandler

st.set_page_config(page_title="Smart Excel Interactive Tool", layout="wide")

st.title("Smart Excel Interactive Tool")
st.markdown("### Interactive Threshold & Coloring Analysis")

# Init Handler
handler = ExcelHandler()

# Init Session State for Overrides
if 'manual_overrides' not in st.session_state:
    st.session_state.manual_overrides = {} # {(row_idx, col_idx): 'HEX'}

# 1. Upload
uploaded_file = st.file_uploader("Upload Excel File", type=["xlsx"])

if uploaded_file:
    # Save to temp for openpyxl processing
    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
        tmp.write(uploaded_file.getvalue())
        tmp_path = tmp.name
    
    # 2. Sheet Selection
    sheet_names = handler.get_sheet_info(tmp_path)
    
    col1, col2 = st.columns([1, 2])
    with col1:
        st.info(f"File contains {len(sheet_names)} sheets.")
        selected_sheet = st.selectbox("Select Sheet to Analyze", sheet_names)
    
    if selected_sheet:
        st.divider()
        st.subheader(f"Analyzing: {selected_sheet}")
        
        # 3. Interactive Viewer (Rows Analysis)
        # Read top rows with styles
        # REMOVED LIMIT: Read all rows to allow full context if needed, though for header detection we might just need top.
        # But user wants full sheet preview.
        rows_data = handler.read_sheet_with_styles(tmp_path, selected_sheet, limit=None)
        
        # Heuristic Detection of Threshold Rows
        detected_rows = handler.detect_threshold_rows(rows_data)
        
        # Visualizer: Custom Table with Colors
        st.write("#### Header & Threshold Inspector")
        st.caption("Inspect the top rows to verify threshold colors. The app attempts to auto-detect colored header rows.")
        
        # Create a display-friendly dataframe for visual checking
        preview_data = []
        for r in rows_data:
            # Convert to string to avoid Arrow/Streamlit type errors with Datetime objects
            row_dict = {f"Col {i}": str(c['value']) if c['value'] is not None else "" for i, c in enumerate(r)}
            preview_data.append(row_dict)
        
        df_preview = pd.DataFrame(preview_data)
        # Sync Index with Excel Row Numbers (1-based)
        df_preview.index = range(1, len(df_preview) + 1)
        
        # Helper for Pandas Styler
        def color_styler(row):
            # row.name is now 1-based Excel Row Number
            # We need 0-based index for rows_data access
            r_idx = row.name - 1
            
            styles = []
            default_style = ""
            
            if 0 <= r_idx < len(rows_data):
                row_meta = rows_data[r_idx]
                for i in range(len(row)):
                    if i < len(row_meta):
                        cell = row_meta[i]
                        bg = cell.get('bg_color')
                        if bg and len(str(bg)) > 1:
                            clean_bg = str(bg).replace('#', '')
                            styles.append(f"background-color: #{clean_bg}; color: black")
                        else:
                             styles.append(default_style)
                    else:
                        styles.append(default_style)
            else:
                 styles = [default_style] * len(row)
                 
            return styles

        # --- User Selection of Threshold Rows ---
        col_sel, col_view = st.columns([1, 3])
        
        with col_sel:
            st.write("##### Configuration")
            # Allow user to override detected rows
            threshold_rows_input = st.multiselect(
                "Select Threshold Rows (Indices)", 
                options=list(range(len(rows_data))),
                default=detected_rows,
                format_func=lambda x: f"Row {x+1}" # User friendly 1-based
            )
            
            # Select Target Columns
            # Improved Naming: Try to find a Header Row (e.g. containing "Well ID" or "Parameter")
            # Scan top 20 rows
            header_row_idx = 0
            df_simple = pd.read_excel(uploaded_file, sheet_name=selected_sheet, header=None, nrows=20)
            
            for i, row in df_simple.iterrows():
                row_str = " ".join([str(x) for x in row if pd.notna(x)])
                if "Well ID" in row_str or "Parameter" in row_str or "pH" in row_str:
                    header_row_idx = i
                    break
            
            # Generate Label from Header Row
            col_options = []
            for i in range(len(df_simple.columns)):
                # Get Name
                name = str(df_simple.iloc[header_row_idx, i])
                if name == "nan": name = f"Col {i}"
                if len(name) > 20: name = name[:17] + "..."
                col_options.append(f"{name} (Idx {i})")
            
            target_cols_indices = st.multiselect(
                "Select Target Columns to Color",
                options=list(range(len(df_simple.columns))),
                format_func=lambda x: col_options[x]
            )
            
            # Clear overrides if sheet changes
            if "last_sheet" not in st.session_state:
                st.session_state.last_sheet = selected_sheet
            elif st.session_state.last_sheet != selected_sheet:
                st.session_state.manual_overrides = {}
                st.session_state.active_row_idx = None # Clear active row too
                st.session_state.last_sheet = selected_sheet
                st.rerun()

            # --- DEBUG: Show Detected Rules ---
            if target_cols_indices and threshold_rows_input:
                with st.expander("🕵️ View Detected Rules (Click to Expand)", expanded=True):
                    st.caption("Verifying exactly what the app sees...")
                    # We need to perform parsing *now* to show this
                    # Reuse internal logic - maybe expose a `get_rules` method? 
                    # Or just quick parse here for display.
                    
                    for col_idx in target_cols_indices:
                        st.markdown(f"**{col_options[col_idx]}**")
                        # Parse
                        # We need the values from rows_data
                        found_any = False
                        for r_idx in threshold_rows_input:
                            if r_idx < len(rows_data):
                                cell_val = rows_data[r_idx][col_idx]['value']
                                bg = rows_data[r_idx][col_idx]['bg_color']
                                if cell_val and bg:
                                    parsed = handler.parse_cell_value(str(cell_val))
                                    if parsed:
                                        found_any = True
                                        # Show nice badge
                                        # Use HTML for color pill
                                        clean_bg = str(bg).replace('#', '')
                                        cond_str = ""
                                        if parsed['type'] == 'range': cond_str = f"{parsed['min']} to {parsed['max']}"
                                        elif parsed['type'] == 'less': cond_str = f"< {parsed['max']}"
                                        elif parsed['type'] == 'greater': cond_str = f"> {parsed['min']}"
                                        elif parsed['type'] == 'implicit': cond_str = f"> {parsed['min']} (Implicit)"
                                        
                                        st.markdown(
                                            f"""<span style='background-color:#{clean_bg}; padding:2px 6px; border-radius:4px; color:black; border:1px solid #ccc'>
                                            Condition: <b>{cond_str}</b>
                                            </span> (Row {r_idx+1}: "{cell_val}")""", 
                                            unsafe_allow_html=True
                                        )
                        if not found_any:
                            st.warning(f"No valid rules found in selected rows for this column.")
            
        with col_view:
            # Display Styled Table
            styled_df = df_preview.style.apply(color_styler, axis=1)
            st.dataframe(styled_df, height=600)
            
            
            # --- PREVIEW SECTION ---
            st.write("##### 5. Preview")
            # --- DEBUG: Inputs ---
            st.write(f"DEBUG CHECKS -> Cols: {target_cols_indices}, Rows: {threshold_rows_input}")
            
            # --- DEBUG STATE (Global) ---
            with st.expander("🛠️ Debug Interactive State (Global)", expanded=True):
                st.write(f"Active Index: {st.session_state.get('active_row_idx')}")
                st.write(f"Overrides Count: {len(st.session_state.manual_overrides)}")
            
            if target_cols_indices and threshold_rows_input:
                try:
                    # --- OPTIMIZATION: Cache Base Preview ---
                    # distinct from overrides. Reruns only if file/sheet/rules change.
                    @st.cache_data(show_spinner=False)
                    def get_cached_preview(f_path, s_name, t_cols, t_rows, h_row):
                        return handler.preview_thresholds(
                            f_path, s_name, t_cols, t_rows, 
                            data_start_row=None, limit=None, header_row_idx=h_row,
                            manual_overrides=None # Do NOT bake overrides into cache
                        )

                    prev_rows, start_row = get_cached_preview(tmp_path, selected_sheet, target_cols_indices, threshold_rows_input, header_row_idx)
                    
                    st.caption(f"Previewing result (Data starts ~Row {start_row}).")
                    
                    # Prepare preview DF
                    prev_data = []
                    
                    # We need to apply overrides HERE, in memory.
                    # prev_rows is 0-based list starting from start_row.
                    # override keys are (row_0_based_absolute, col_idx).
                    # Map: absolute_row = start_row + rel_idx.
                    
                    for rel_idx, r in enumerate(prev_rows):
                        current_abs_row_idx = start_row + rel_idx - 1 # wait, start_row is 1-based index? 
                        # In preview_thresholds:
                        # existing_hex = self.get_cell_hex(cell)
                        # cell.row is 1-based.
                        # current_row_idx = cell.row - 1. (0-based)
                        # preview_rows starts iterating at final_start_row (1-based implied if passed to iter_rows?)
                        # ws.iter_rows(min_row=final_start_row)
                        # First row yielded has cell.row = final_start_row.
                        # So first item in prev_rows corresponds to 0-based index: final_start_row - 1.
                        
                        curr_zero_based_row_idx = (start_row - 1) + rel_idx
                        
                        d = {}
                        for k, v in r.items():
                             # k is "Col 0", "Col 1"...
                             col_idx = int(k.split(' ')[1])
                             val = v['value']
                             bg = v['bg']
                             
                             # --- Apply Override In-Memory ---
                             if (curr_zero_based_row_idx, col_idx) in st.session_state.manual_overrides:
                                 ov_hex = st.session_state.manual_overrides[(curr_zero_based_row_idx, col_idx)]
                                 if ov_hex:
                                     bg = ov_hex
                                 else:
                                     bg = None # Explicit Clear
                             # --------------------------------
                             
                             d[k] = val
                             # Store BG separately? 
                             # The DF only stores values. The STYLER applies colors.
                             # We need to pass color info to styler. 
                             # Styler takes function: row -> styles.
                             # Row contains values. 
                             # We can use a parallel Dataframe for colors? Or stick to lookup.
                        prev_data.append(d)
                        
                    df_prev_show = pd.DataFrame(prev_data)
                    
                    # Set Index to match Excel Row Numbers
                    df_prev_show.index = range(start_row, start_row + len(df_prev_show))
                    
                    def preview_styler(row):
                        list_idx = row.name - start_row
                        
                        styles = []
                        if 0 <= list_idx < len(prev_rows):
                            # Base Colors
                            row_meta = prev_rows[list_idx]
                            
                            for col_name in df_prev_show.columns: 
                                c_idx = int(col_name.split(' ')[1])
                                
                                # Resolve Color (Base + Override)
                                # Base
                                cell_meta = row_meta.get(col_name)
                                bg = cell_meta.get('bg') if cell_meta else None
                                
                                # Override check (Fast lookup)
                                # row.name is 1-based Excel Index
                                zero_based = row.name - 1
                                if (zero_based, c_idx) in st.session_state.manual_overrides:
                                    ov = st.session_state.manual_overrides[(zero_based, c_idx)]
                                    bg = ov if ov else None
                                
                                if bg:
                                     styles.append(f"background-color: #{bg}; color: black")
                                else:
                                     styles.append("")
                        else:
                             styles = [""] * len(row)
                        return styles
                    
                    selection = st.dataframe(
                        df_prev_show.style.apply(preview_styler, axis=1), 
                        height=500,
                        on_select="rerun", # Enable Selection
                        selection_mode="single-row"
                    )
                    
                    # --- INTERACTIVE CORRECTION ---
                    # --- DEBUG STATE ---
                    with st.expander("🛠️ Debug Interactive State", expanded=False):
                        st.write(f"Selection Object Type: {type(selection)}")
                        #st.write(f"Selection Object Content: {selection}")
                        st.write(f"Active Index: {st.session_state.get('active_row_idx')}")
                        st.write(f"Overrides Count: {len(st.session_state.manual_overrides)}")
                        st.write(st.session_state.manual_overrides)

                    if selection and "selection" in selection and "rows" in selection["selection"]:
                        rows_sel = selection["selection"]["rows"]
                        if rows_sel:
                            row_pos = rows_sel[0]
                            if row_pos < len(df_prev_show):
                                excel_row_num = df_prev_show.index[row_pos]
                                # Save to Session State
                                st.session_state.active_row_idx = excel_row_num
                                st.rerun() # Force rerun to update UI immediately for new row selection?
                                # Actually, better not to rerun inside render unless needed. 
                                # But here we want the Panel to show immediately.
                                # The Panel logic below checks 'active_row_idx'. 
                                # Since we just set it, it WILL show in this run. No rerun needed.

                    # --- INTERACTIVE CORRECTION PANEL ---
                    # Always show if we have an active row in state
                    if 'active_row_idx' in st.session_state:
                         excel_row_num = st.session_state.active_row_idx
                         logic_row_idx = excel_row_num - 1
                         
                         st.info(f"Editing Row: {excel_row_num}")
                         
                         # Add Close Button
                         if st.button("Close / Deselect Row", key="close_row_edit"):
                             del st.session_state.active_row_idx
                             st.rerun()
                         
                         cols = st.columns(len(target_cols_indices))
                         for i, col_idx in enumerate(target_cols_indices):
                            c_name = col_options[col_idx]
                            
                            # Determine Current Color state
                            curr_bg = None
                            if (logic_row_idx, col_idx) in st.session_state.manual_overrides:
                                curr_bg = st.session_state.manual_overrides[(logic_row_idx, col_idx)]
                            else:
                                rel_idx = excel_row_num - start_row
                                if 0 <= rel_idx < len(prev_rows):
                                     cell_data = prev_rows[rel_idx]
                                     curr_bg = cell_data.get(f"Col {col_idx}", {}).get('bg')

                            # Display Status
                            show_color = curr_bg if curr_bg else "FFFFFF"
                            lbl = f"{c_name}"
                            
                            with cols[i]:
                                st.caption(lbl)
                                cycle_colors = [None, "FF6B6B", "4CAF50", "FFD700"]
                                
                                try:
                                    curr_idx = cycle_colors.index(curr_bg)
                                except:
                                    curr_idx = 0
                                    
                                next_idx = (curr_idx + 1) % len(cycle_colors)
                                next_color = cycle_colors[next_idx]
                                
                                st.markdown(
                                    f"<div style='width:100%; height:30px; background-color:#{show_color}; border:1px solid #ccc; margin-bottom:5px;'></div>", 
                                    unsafe_allow_html=True
                                )
                                
                                if st.button(f"Cycle Color", key=f"btn_{logic_row_idx}_{col_idx}"):
                                    st.session_state.manual_overrides[(logic_row_idx, col_idx)] = next_color
                                    st.rerun()

                except Exception as e:
                    st.error(f"Could not generate preview: {e}")
                    st.exception(e)
            
            # 6. Action: Apply
            if st.button("Apply Threshold Coloring", type="primary"):
                if not target_cols_indices:
                    st.warning("Please select at least one column.")
                elif not threshold_rows_input:
                    st.warning("Please select rows.")
                else:
                    with st.spinner("Processing..."):
                        out_path = f"processed_{selected_sheet}.xlsx"
                        handler.apply_thresholds(
                            tmp_path, 
                            selected_sheet, 
                            target_cols_indices, 
                            threshold_rows_input, 
                            out_path,
                            data_start_row=None,
                            header_row_idx=header_row_idx,
                            manual_overrides=st.session_state.manual_overrides
                        )
                        st.success("Result Ready!")
                        with open(out_path, "rb") as f:
                            st.download_button(
                                "📥 Download Colored Excel", 
                                f, 
                                file_name=f"colored_{selected_sheet}.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                            )
