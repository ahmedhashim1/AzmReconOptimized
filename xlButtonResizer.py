import win32com.client
import pythoncom


def resize_excel_buttons(file_name, sheet_name, default_width=100, default_height=30, align_column='S'):
    """
    Resize all button controls in an Excel sheet to default dimensions and align them under a specific column.
    Checks for: Form Controls, ActiveX Controls, Shapes, and all button types.

    Parameters:
    -----------
    file_name : str
        Name of the Excel file (must be already open in Excel)
    sheet_name : str
        Name of the sheet containing the buttons
    default_width : float, optional
        Default width for buttons in points (default: 100)
    default_height : float, optional
        Default height for buttons in points (default: 30)
    align_column : str, optional
        Column letter to align buttons under (default: 'S')

    Returns:
    --------
    dict
        Summary of buttons found and resized
    """
    pythoncom.CoInitialize()

    try:
        # Connect to running Excel instance
        excel = win32com.client.GetActiveObject("Excel.Application")

        # Find the workbook by name
        workbook = None
        for wb in excel.Workbooks:
            if wb.Name == file_name or wb.FullName.endswith(file_name):
                workbook = wb
                break

        if not workbook:
            raise ValueError(f"Workbook '{file_name}' not found in open Excel files")

        # Get the specified sheet
        try:
            sheet = workbook.Sheets(sheet_name)
        except:
            raise ValueError(f"Sheet '{sheet_name}' not found in workbook '{file_name}'")

        results = {
            'form_buttons_found': 0,
            'form_buttons_resized': 0,
            'activex_buttons_found': 0,
            'activex_buttons_resized': 0,
            'other_buttons_found': 0,
            'other_buttons_resized': 0,
            'buttons_moved': 0,
            'details': []
        }

        # Get the left position of the target column
        column_number = ord(align_column.upper()) - ord('A') + 1
        target_left = sheet.Columns(column_number).Left

        print(f"\n{'=' * 80}")
        print(f"Scanning sheet '{sheet_name}' in workbook '{file_name}'")
        print(f"Target column for alignment: {align_column} (Left position: {target_left:.1f})")
        print(f"{'=' * 80}")

        # Method 1: Check all Shapes (covers Form Controls and some buttons)
        print(f"\n[Method 1] Checking Shapes collection...")
        print(f"Total shapes found: {sheet.Shapes.Count}")

        for shape in sheet.Shapes:
            try:
                shape_type = shape.Type
                shape_name = shape.Name

                print(f"\n  Shape: {shape_name}")
                print(f"    Type Code: {shape_type}")

                # Check various button types
                is_button = False
                button_type = "Unknown"

                # Type 8 = msoFormControl
                if shape_type == 8:
                    try:
                        control_type = shape.FormControlType
                        print(f"    Form Control Type: {control_type}")

                        # xlButtonControl = 8, xlDropDown = 1, etc.
                        if control_type == 8:
                            is_button = True
                            button_type = "Form Control Button (xlButtonControl)"
                            results['form_buttons_found'] += 1
                    except Exception as e:
                        print(f"    Could not determine FormControlType: {e}")

                # Type 1 = msoAutoShape (can include button shapes)
                elif shape_type == 1:
                    button_type = "AutoShape (potential button)"
                    is_button = True
                    results['other_buttons_found'] += 1

                # Type 12 = msoOLEControlObject (ActiveX)
                elif shape_type == 12:
                    button_type = "OLE Control Object (ActiveX)"
                    is_button = True
                    results['activex_buttons_found'] += 1

                # Type 13 = msoPicture
                elif shape_type == 13:
                    # Check if picture has macro assigned (used as button)
                    try:
                        if hasattr(shape, 'OnAction') and shape.OnAction:
                            button_type = "Picture with Macro (Button)"
                            is_button = True
                            results['other_buttons_found'] += 1
                    except:
                        pass

                # Type 17 = msoTextBox
                elif shape_type == 17:
                    # Check if textbox has macro assigned (used as button)
                    try:
                        if hasattr(shape, 'OnAction') and shape.OnAction:
                            button_type = "TextBox with Macro (Button)"
                            is_button = True
                            results['other_buttons_found'] += 1
                    except:
                        pass

                # Any shape with OnAction can be a button
                else:
                    try:
                        if hasattr(shape, 'OnAction') and shape.OnAction:
                            button_type = f"Shape Type {shape_type} with Macro (Button)"
                            is_button = True
                            results['other_buttons_found'] += 1
                    except:
                        pass

                if is_button:
                    current_width = shape.Width
                    current_height = shape.Height
                    current_left = shape.Left

                    print(f"    ✓ BUTTON FOUND: {button_type}")
                    print(f"    Current size: {current_width:.1f} x {current_height:.1f}")
                    print(f"    Current left position: {current_left:.1f}")

                    # Check if resize is needed
                    needs_resize = abs(current_width - default_width) > 0.1 or abs(
                        current_height - default_height) > 0.1
                    needs_move = abs(current_left - target_left) > 0.1

                    if needs_resize:
                        shape.Width = default_width
                        shape.Height = default_height

                        if 'form' in button_type.lower():
                            results['form_buttons_resized'] += 1
                        elif 'activex' in button_type.lower():
                            results['activex_buttons_resized'] += 1
                        else:
                            results['other_buttons_resized'] += 1

                        print(f"    New size: {shape.Width:.1f} x {shape.Height:.1f}")

                    if needs_move:
                        shape.Left = target_left
                        results['buttons_moved'] += 1
                        print(f"    New left position: {shape.Left:.1f} (moved to column {align_column})")

                    if needs_resize or needs_move:
                        status = "✓ RESIZED" if needs_resize and needs_move else (
                            "✓ RESIZED" if needs_resize else "✓ MOVED")
                    else:
                        status = "Already correct size and position"

                    results['details'].append({
                        'name': shape_name,
                        'type': button_type,
                        'old_size': f"{current_width:.1f} x {current_height:.1f}",
                        'new_size': f"{shape.Width:.1f} x {shape.Height:.1f}",
                        'old_left': f"{current_left:.1f}",
                        'new_left': f"{shape.Left:.1f}",
                        'status': status
                    })
                    print(f"    Status: {status}")

            except Exception as e:
                print(f"    Error processing shape: {str(e)}")

        # Method 2: Check OLEObjects (ActiveX Controls)
        print(f"\n[Method 2] Checking OLEObjects collection...")
        try:
            ole_count = sheet.OLEObjects().Count
            print(f"Total OLE objects found: {ole_count}")

            for i in range(1, ole_count + 1):
                try:
                    ole_obj = sheet.OLEObjects(i)
                    print(f"\n  OLE Object: {ole_obj.Name}")
                    print(f"    ProgID: {ole_obj.progID}")

                    # Check for various button types
                    is_button = False
                    button_type = "Unknown ActiveX"

                    prog_id_lower = ole_obj.progID.lower()

                    if "commandbutton" in prog_id_lower or "button" in prog_id_lower:
                        is_button = True
                        button_type = "ActiveX CommandButton"
                    elif "togglebutton" in prog_id_lower:
                        is_button = True
                        button_type = "ActiveX ToggleButton"
                    elif "forms." in prog_id_lower:
                        is_button = True
                        button_type = f"ActiveX Forms Control ({ole_obj.progID})"

                    if is_button:
                        current_width = ole_obj.Width
                        current_height = ole_obj.Height
                        current_left = ole_obj.Left

                        print(f"    ✓ BUTTON FOUND: {button_type}")
                        print(f"    Current size: {current_width:.1f} x {current_height:.1f}")
                        print(f"    Current left position: {current_left:.1f}")

                        # Check if already counted in shapes
                        already_counted = any(d['name'] == ole_obj.Name for d in results['details'])

                        if not already_counted:
                            results['activex_buttons_found'] += 1

                            needs_resize = abs(current_width - default_width) > 0.1 or abs(
                                current_height - default_height) > 0.1
                            needs_move = abs(current_left - target_left) > 0.1

                            if needs_resize:
                                ole_obj.Width = default_width
                                ole_obj.Height = default_height
                                results['activex_buttons_resized'] += 1
                                print(f"    New size: {ole_obj.Width:.1f} x {ole_obj.Height:.1f}")

                            if needs_move:
                                ole_obj.Left = target_left
                                results['buttons_moved'] += 1
                                print(f"    New left position: {ole_obj.Left:.1f} (moved to column {align_column})")

                            if needs_resize or needs_move:
                                status = "✓ RESIZED" if needs_resize and needs_move else (
                                    "✓ RESIZED" if needs_resize else "✓ MOVED")
                            else:
                                status = "Already correct size and position"

                            results['details'].append({
                                'name': ole_obj.Name,
                                'type': button_type,
                                'old_size': f"{current_width:.1f} x {current_height:.1f}",
                                'new_size': f"{ole_obj.Width:.1f} x {ole_obj.Height:.1f}",
                                'old_left': f"{current_left:.1f}",
                                'new_left': f"{ole_obj.Left:.1f}",
                                'status': status
                            })
                            print(f"    Status: {status}")
                        else:
                            print(f"    (Already processed in Shapes collection)")

                except Exception as e:
                    print(f"    Error processing OLE object: {str(e)}")

        except Exception as e:
            print(f"Error accessing OLEObjects: {str(e)}")

        # Method 3: Check Buttons collection (if it exists)
        print(f"\n[Method 3] Checking Buttons collection...")
        try:
            buttons = sheet.Buttons()
            button_count = buttons.Count
            print(f"Total buttons in Buttons collection: {button_count}")

            for i in range(1, button_count + 1):
                try:
                    btn = buttons.Item(i)
                    print(f"\n  Button: {btn.Name}")

                    # Check if already counted
                    already_counted = any(d['name'] == btn.Name for d in results['details'])

                    if not already_counted:
                        current_width = btn.Width
                        current_height = btn.Height
                        current_left = btn.Left

                        print(f"    ✓ BUTTON FOUND: Form Button")
                        print(f"    Current size: {current_width:.1f} x {current_height:.1f}")
                        print(f"    Current left position: {current_left:.1f}")

                        results['form_buttons_found'] += 1

                        needs_resize = abs(current_width - default_width) > 0.1 or abs(
                            current_height - default_height) > 0.1
                        needs_move = abs(current_left - target_left) > 0.1

                        if needs_resize:
                            btn.Width = default_width
                            btn.Height = default_height
                            results['form_buttons_resized'] += 1
                            print(f"    New size: {btn.Width:.1f} x {btn.Height:.1f}")

                        if needs_move:
                            btn.Left = target_left
                            results['buttons_moved'] += 1
                            print(f"    New left position: {btn.Left:.1f} (moved to column {align_column})")

                        if needs_resize or needs_move:
                            status = "✓ RESIZED" if needs_resize and needs_move else (
                                "✓ RESIZED" if needs_resize else "✓ MOVED")
                        else:
                            status = "Already correct size and position"

                        results['details'].append({
                            'name': btn.Name,
                            'type': "Form Button",
                            'old_size': f"{current_width:.1f} x {current_height:.1f}",
                            'new_size': f"{btn.Width:.1f} x {btn.Height:.1f}",
                            'old_left': f"{current_left:.1f}",
                            'new_left': f"{btn.Left:.1f}",
                            'status': status
                        })
                        print(f"    Status: {status}")
                    else:
                        print(f"    (Already processed)")

                except Exception as e:
                    print(f"    Error processing button: {str(e)}")

        except Exception as e:
            print(f"Error accessing Buttons collection: {str(e)}")

        # Print summary
        total_found = results['form_buttons_found'] + results['activex_buttons_found'] + results['other_buttons_found']
        total_resized = results['form_buttons_resized'] + results['activex_buttons_resized'] + results[
            'other_buttons_resized']

        print("\n" + "=" * 80)
        print("SUMMARY")
        print("=" * 80)
        print(f"Form Control Buttons: {results['form_buttons_found']} found, {results['form_buttons_resized']} resized")
        print(
            f"ActiveX Buttons: {results['activex_buttons_found']} found, {results['activex_buttons_resized']} resized")
        print(f"Other Button Types: {results['other_buttons_found']} found, {results['other_buttons_resized']} resized")
        print(
            f"\nTOTAL: {total_found} buttons found, {total_resized} resized, {results['buttons_moved']} moved to column {align_column}")
        print("=" * 80)

        if total_found == 0:
            print("\n⚠ WARNING: No buttons were found!")
            print("Please verify:")
            print(f"  1. The workbook name is correct: '{file_name}'")
            print(f"  2. The sheet name is correct: '{sheet_name}'")
            print(f"  3. The Excel file is currently open")
            print(f"  4. The sheet actually contains button controls")

        return results

    except Exception as e:
        print(f"\n❌ ERROR: {str(e)}")
        import traceback
        traceback.print_exc()
        raise
    finally:
        pythoncom.CoUninitialize()


# Example usage
if __name__ == "__main__":
    # Example: Resize buttons in an already open Excel file
    file_name = "All Billers Reconciliation Summary - October.xlsm"  # Name of the open workbook (just filename or full path)
    sheet_name = "23-Oct"  # Name of the sheet

    # Optional: Set custom default dimensions (in points)
    default_width = 100  # points
    default_height = 30  # points
    align_column = 'S'  # Column to align buttons under

    try:
        results = resize_excel_buttons(
            file_name=file_name,
            sheet_name=sheet_name,
            default_width=default_width,
            default_height=default_height,
            align_column=align_column
        )

        print("\n✓ Operation completed successfully!")

    except Exception as e:
        print(f"\n❌ Failed to resize buttons: {str(e)}")