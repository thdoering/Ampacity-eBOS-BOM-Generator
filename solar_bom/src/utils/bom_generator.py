import pandas as pd
import os
from typing import Dict, List, Any, Optional
from ..models.block import BlockConfig, WiringType
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter


class BOMGenerator:
    """Utility class for generating Bill of Materials from block configurations"""
    
    # Constants for BOM calculations
    CABLE_WASTE_FACTOR = 1.05  # 5% extra for waste/installation
    
    def __init__(self, blocks: Dict[str, BlockConfig]):
        """
        Initialize BOM Generator with block configurations
        
        Args:
            blocks: Dictionary of block configurations (id -> BlockConfig)
        """
        self.blocks = blocks
    
    def calculate_cable_quantities(self) -> Dict[str, Dict[str, Any]]:
        """
        Calculate cable quantities by block and type
        
        Returns:
            Dictionary with quantities by block and component type
            {
                'block_id': {
                    'component_type': {
                        'description': description,
                        'quantity': quantity,
                        'unit': unit,
                        'category': category  # Added to categorize components
                    }
                }
            }
        """
        quantities = {}
        
        for block_id, block in self.blocks.items():
            if not block.wiring_config:
                continue
                
            block_quantities = {}
            cable_lengths = block.calculate_cable_lengths()
            
            # Get tracker counts categorized by strings per tracker
            tracker_counts = {}
            for pos in block.tracker_positions:
                if pos.template:
                    strings_per_tracker = pos.template.strings_per_tracker
                    key = f"{strings_per_tracker}-String Tracker"
                    if key not in tracker_counts:
                        tracker_counts[key] = 0
                    tracker_counts[key] += 1
            
            # Add trackers to quantities (as structural components)
            for tracker_type, count in tracker_counts.items():
                block_quantities[tracker_type] = {
                    'description': f"{tracker_type}",
                    'quantity': count,
                    'unit': 'units',
                    'category': 'Structural'
                }
            
            if block.wiring_config.wiring_type == WiringType.HOMERUN:
                # For homerun, we track string cable length
                if 'string_cable' in cable_lengths:
                    length = cable_lengths['string_cable'] * self.CABLE_WASTE_FACTOR
                    # Convert from meters to feet (1m = 3.28084ft)
                    length_feet = round(length * 3.28084, 1)
                    
                    # Get the cable size for more specific description
                    cable_size = block.wiring_config.string_cable_size
                    
                    block_quantities[f'String Wire ({cable_size})'] = {
                        'description': f'DC String Wire {cable_size}',
                        'quantity': length_feet,
                        'unit': 'feet',
                        'category': 'eBOS'
                    }
                    
                    # Get the whip cable size
                    whip_cable_size = getattr(block.wiring_config, 'whip_cable_size', '6 AWG')  # Default to 6 AWG if not specified

                    # Calculate whip length - estimate 2 whips per tracker with 3m (10ft) per whip
                    whip_count = len([pos for pos in block.tracker_positions])
                    whip_length_per_tracker = 6  # 3m per whip x 2 whips = 6m
                    whip_length_feet = round(whip_count * whip_length_per_tracker * 3.28084 * self.CABLE_WASTE_FACTOR, 1)

                    block_quantities[f'Whip Cable ({whip_cable_size})'] = {
                        'description': f'DC Whip Cable {whip_cable_size}',
                        'quantity': whip_length_feet,
                        'unit': 'feet',
                        'category': 'eBOS'
                    }
                    
            else:  # HARNESS configuration
                # Add harnesses by number of strings they connect
                harness_counts = self._count_harnesses_by_size(block)
                
                # Add each harness type
                for string_count, count in harness_counts.items():
                    harness_cable_size = block.wiring_config.harness_cable_size
                    string_cable_size = block.wiring_config.string_cable_size
                    
                    block_quantities[f'{string_count}-String Harness'] = {
                        'description': f'{string_count}-String Harness ({harness_cable_size} trunk, {string_cable_size} drops)',
                        'quantity': count,
                        'unit': 'units',
                        'category': 'eBOS'
                    }
                
                # Add trunk wire from harness to inverter (whips)
                # Note: String cables are part of the harness in this configuration
                whip_count = len([pos for pos in block.tracker_positions])
                harness_cable_size = block.wiring_config.harness_cable_size
                
                # Add whip cables 
                whip_cable_size = getattr(block.wiring_config, 'whip_cable_size', '6 AWG')  # Default to 6 AWG if not specified
                # Calculate whip length - estimate 2 whips per tracker with 3m (10ft) per whip
                whip_count = len([pos for pos in block.tracker_positions])
                whip_length_per_tracker = 6  # 3m per whip x 2 whips = 6m
                whip_length_feet = round(whip_count * whip_length_per_tracker * 3.28084 * self.CABLE_WASTE_FACTOR, 1)

                block_quantities[f'Whip Cable ({whip_cable_size})'] = {
                    'description': f'DC Whip Cable {whip_cable_size}',
                    'quantity': whip_length_feet,
                    'unit': 'feet',
                    'category': 'eBOS'
                }
            
            quantities[block_id] = block_quantities
        
        return quantities
        
    def _count_harnesses_by_size(self, block: BlockConfig) -> Dict[int, int]:
        """
        Count harnesses by number of strings they connect
        
        Args:
            block: Block configuration
            
        Returns:
            Dictionary mapping string count to harness count
        """
        harness_counts = {}
        
        if not block.wiring_config or block.wiring_config.wiring_type != WiringType.HARNESS:
            return harness_counts
            
        # Count strings per tracker
        for pos in block.tracker_positions:
            if pos.template:
                string_count = len(pos.strings)
                if string_count not in harness_counts:
                    harness_counts[string_count] = 0
                harness_counts[string_count] += 1
                
        return harness_counts
    
    def generate_summary_data(self, quantities: Dict[str, Dict[str, Any]]) -> pd.DataFrame:
        """
        Generate summary data for BOM with categories
        
        Args:
            quantities: Component quantities by block
            
        Returns:
            DataFrame with summary data
        """
        # Initialize dictionaries to track totals by component, description, and category
        component_totals = {}
        
        # Sum up quantities for each component type across all blocks
        for block_id, block_quantities in quantities.items():
            for component_type, details in block_quantities.items():
                description = details['description']
                quantity = details['quantity']
                unit = details['unit']
                category = details['category']
                
                key = (component_type, description, unit, category)
                if key not in component_totals:
                    component_totals[key] = 0
                component_totals[key] += quantity
        
        # Convert to DataFrame
        summary_data = []
        for (component_type, description, unit, category), quantity in component_totals.items():
            summary_data.append({
                'Category': category,
                'Component Type': component_type,
                'Description': description,
                'Quantity': round(quantity, 1) if unit == 'feet' else int(quantity),
                'Unit': unit
            })
        
        # Sort by category then component type
        summary_data = sorted(summary_data, key=lambda x: (x['Category'], x['Component Type']))
        
        return pd.DataFrame(summary_data)
    
    def generate_detailed_data(self, quantities: Dict[str, Dict[str, Any]]) -> pd.DataFrame:
        """
        Generate detailed data for BOM
        
        Args:
            quantities: Component quantities by block
            
        Returns:
            DataFrame with detailed data
        """
        detailed_data = []
        
        for block_id, block_quantities in quantities.items():
            for component_type, details in block_quantities.items():
                description = details['description']
                quantity = details['quantity']
                unit = details['unit']
                category = details['category']
                
                detailed_data.append({
                    'Block': block_id,
                    'Category': category,
                    'Component Type': component_type,
                    'Description': description,
                    'Quantity': round(quantity, 1) if unit == 'feet' else int(quantity),
                    'Unit': unit
                })
        
        # Sort by block ID, category, and component type
        detailed_data = sorted(detailed_data, key=lambda x: (x['Block'], x['Category'], x['Component Type']))
        
        return pd.DataFrame(detailed_data)

    
    def export_bom_to_excel(self, filepath: str, project_info: Optional[Dict[str, Any]] = None) -> bool:
        """
        Export BOM to Excel file
        
        Args:
            filepath: Path to save the Excel file
            project_info: Optional dictionary with project information
            
        Returns:
            True if export successful, False otherwise
        """
        writer = None
        try:
            # Calculate quantities
            quantities = self.calculate_cable_quantities()
            
            # Generate summary and detailed data
            summary_data = self.generate_summary_data(quantities)
            detailed_data = self.generate_detailed_data(quantities)
            
            # Generate project info if not provided
            if project_info is None:
                project_info = self.generate_project_info()
            
            # Create Excel writer
            writer = pd.ExcelWriter(filepath, engine='openpyxl')
            
            # Write summary data
            summary_data.to_excel(writer, sheet_name='BOM Summary', index=False, startrow=10)  # Start after project info
            
            # Write detailed data
            detailed_data.to_excel(writer, sheet_name='Block Details', index=False)
            
            # Get workbook
            workbook = writer.book
            
            # Get summary worksheet
            summary_sheet = writer.sheets['BOM Summary']
            
            # Add project info to summary sheet
            row = 1
            summary_sheet.merge_cells(f'A{row}:E{row}')
            summary_sheet.cell(row=row, column=1, value="Project Information").font = Font(bold=True, size=14)
            
            if project_info:
                for i, (key, value) in enumerate(project_info.items()):
                    # Format system size to 2 decimal places
                    if key == 'System Size (kW DC)' and isinstance(value, (int, float)):
                        value = round(value, 2)
                        
                    row = i + 2
                    summary_sheet.cell(row=row, column=1, value=key).font = Font(bold=True)
                    summary_sheet.cell(row=row, column=2, value=value)
            else:
                # Generate project info from blocks if not provided
                generated_info = self.generate_project_info()
                for i, (key, value) in enumerate(generated_info.items()):
                    # Format system size to 2 decimal places
                    if key == 'System Size (kW DC)' and isinstance(value, (int, float)):
                        value = round(value, 2)
                        
                    row = i + 2
                    summary_sheet.cell(row=row, column=1, value=key).font = Font(bold=True)
                    summary_sheet.cell(row=row, column=2, value=str(value))  
            
            # Add section header for BOM
            row = 9
            summary_sheet.merge_cells(f'A{row}:E{row}')
            summary_sheet.cell(row=row, column=1, value="Bill of Materials").font = Font(bold=True, size=14)
            
            # Format summary sheet
            self._format_excel_sheet(workbook['BOM Summary'], summary_data, start_row=10)
            
            # Format detailed sheet
            self._format_excel_sheet(workbook['Block Details'], detailed_data)
            
            # Add filter to summary sheet for easier sorting
            summary_sheet.auto_filter.ref = f"A10:E{10 + len(summary_data)}"
            
            # Save the Excel file
            writer.close()
            writer = None  # Set to None after closing to avoid double close in finally block
            
            return True
        except Exception as e:
            print(f"Error exporting BOM: {str(e)}")
            # Re-raise permission errors to be caught by the calling function
            if isinstance(e, PermissionError):
                raise
            return False
        finally:
            # Make sure the writer is closed
            if writer is not None:
                try:
                    writer.close()
                except:
                    pass

    
    def _format_excel_sheet(self, worksheet, data: pd.DataFrame, start_row: int = 1):
        """
        Format Excel worksheet
        
        Args:
            worksheet: openpyxl worksheet
            data: DataFrame with data
            start_row: Row to start formatting from (default=1)
        """
        # Define styles
        header_font = Font(bold=True, color="FFFFFF")
        header_fill = PatternFill(start_color="4F81BD", end_color="4F81BD", fill_type="solid")
        centered_alignment = Alignment(horizontal='center')
        border = Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin')
        )
        
        # Category styles
        category_fill = PatternFill(start_color="E6E6E6", end_color="E6E6E6", fill_type="solid")
        
        # Format headers
        for col_num, column_title in enumerate(data.columns, 1):
            cell = worksheet.cell(row=start_row, column=col_num)
            cell.value = column_title
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = centered_alignment
            cell.border = border
        
        # Format data cells and apply auto-width
        current_category = None
        
        for i, row in enumerate(worksheet.iter_rows(min_row=start_row+1, 
                                                    max_row=start_row+len(data), 
                                                    min_col=1, 
                                                    max_col=len(data.columns)), 1):
            # Check if category has changed
            if 'Category' in data.columns:
                row_index = start_row + i
                category = worksheet.cell(row=row_index, column=1).value
                
                if category != current_category:
                    current_category = category
                    # Add subtle highlight for category rows
                    for cell in row:
                        cell.fill = category_fill
            
            # Add borders to all cells
            for cell in row:
                cell.border = border
        
        # Auto-adjust column width
        for column in worksheet.columns:
            max_length = 0
            column_name = get_column_letter(column[0].column)
            for cell in column:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            adjusted_width = (max_length + 2)
            worksheet.column_dimensions[column_name].width = adjusted_width

    def generate_project_info(self) -> Dict[str, Any]:
        """
        Generate general project information for BOM header
        
        Returns:
            Dictionary with project info
        """
        # Initialize with default values
        info = {
            'System Size (kW DC)': 0,
            'Number of Modules': 0,
            'Module Manufacturer': 'Unknown',
            'Module Model': 'Unknown',
            'DC Collection': 'Unknown',
            'Inverter Manufacturer': 'Unknown',
            'Inverter Model': 'Unknown'
        }
        
        # Calculate total system size and module count
        total_modules = 0
        module_manufacturer = set()
        module_model = set()
        inverter_manufacturer = set()
        inverter_model = set()
        dc_collection_types = set()
        
        for block_id, block in self.blocks.items():
            # Count modules
            if block.tracker_template and block.tracker_template.module_spec:
                module_spec = block.tracker_template.module_spec
                modules_per_tracker = block.tracker_template.get_total_modules()
                tracker_count = len(block.tracker_positions)
                block_modules = modules_per_tracker * tracker_count
                
                total_modules += block_modules
                
                # Add module info
                module_manufacturer.add(module_spec.manufacturer)
                module_model.add(module_spec.model)
                
                # Calculate system size
                info['System Size (kW DC)'] += (block_modules * module_spec.wattage) / 1000
            
            # Add inverter info
            if block.inverter:
                inverter_manufacturer.add(block.inverter.manufacturer)
                inverter_model.add(block.inverter.model)
            
            # Add DC collection type
            if block.wiring_config:
                dc_collection_types.add(block.wiring_config.wiring_type.value)
        
        # Update info dict
        info['Number of Modules'] = total_modules
        info['Module Manufacturer'] = ', '.join(module_manufacturer) if module_manufacturer else 'Unknown'
        info['Module Model'] = ', '.join(module_model) if module_model else 'Unknown'
        info['Inverter Manufacturer'] = ', '.join(inverter_manufacturer) if inverter_manufacturer else 'Unknown'
        info['Inverter Model'] = ', '.join(inverter_model) if inverter_model else 'Unknown'
        info['DC Collection'] = ', '.join(dc_collection_types) if dc_collection_types else 'Unknown'
        
        return info