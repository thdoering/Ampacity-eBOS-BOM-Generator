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
    
    def calculate_cable_quantities(self) -> Dict[str, Dict[str, float]]:
        """
        Calculate cable quantities by block and type
        
        Returns:
            Dictionary of cable quantities by block and type
            {
                'block_id': {
                    'cable_type': length_in_meters
                }
            }
        """
        quantities = {}
        
        for block_id, block in self.blocks.items():
            if not block.wiring_config:
                continue
                
            block_quantities = {}
            cable_lengths = block.calculate_cable_lengths()
            
            if block.wiring_config.wiring_type == WiringType.HOMERUN:
                # For homerun, we use just string cables
                if 'string_cable' in cable_lengths:
                    block_quantities[block.wiring_config.string_cable_size] = cable_lengths['string_cable'] * self.CABLE_WASTE_FACTOR
            else:
                # For harness, we have both string and harness cables
                if 'string_cable' in cable_lengths:
                    block_quantities[block.wiring_config.string_cable_size] = cable_lengths.get('string_cable', 0) * self.CABLE_WASTE_FACTOR
                if 'harness_cable' in cable_lengths:
                    block_quantities[block.wiring_config.harness_cable_size] = cable_lengths.get('harness_cable', 0) * self.CABLE_WASTE_FACTOR
            
            quantities[block_id] = block_quantities
        
        return quantities
    
    def generate_summary_data(self, quantities: Dict[str, Dict[str, float]]) -> pd.DataFrame:
        """
        Generate summary data for BOM
        
        Args:
            quantities: Cable quantities by block and type
            
        Returns:
            DataFrame with summary data
        """
        # Initialize dictionary to store totals by cable type
        cable_totals = {}
        
        # Sum up quantities for each cable type across all blocks
        for block_id, block_quantities in quantities.items():
            for cable_type, length in block_quantities.items():
                if cable_type not in cable_totals:
                    cable_totals[cable_type] = 0
                cable_totals[cable_type] += length
        
        # Convert to DataFrame
        summary_data = []
        for cable_type, length in cable_totals.items():
            summary_data.append({
                'Component Type': 'DC Cable',
                'Description': cable_type,
                'Quantity': round(length, 1),
                'Unit': 'meters'
            })
        
        return pd.DataFrame(summary_data)
    
    def generate_detailed_data(self, quantities: Dict[str, Dict[str, float]]) -> pd.DataFrame:
        """
        Generate detailed data for BOM
        
        Args:
            quantities: Cable quantities by block and type
            
        Returns:
            DataFrame with detailed data
        """
        detailed_data = []
        
        for block_id, block_quantities in quantities.items():
            for cable_type, length in block_quantities.items():
                detailed_data.append({
                    'Block': block_id,
                    'Component Type': 'DC Cable',
                    'Description': cable_type,
                    'Quantity': round(length, 1),
                    'Unit': 'meters'
                })
        
        return pd.DataFrame(detailed_data)
    
    def export_bom_to_excel(self, filepath: str) -> bool:
        """
        Export BOM to Excel file
        
        Args:
            filepath: Path to save the Excel file
            
        Returns:
            True if export successful, False otherwise
        """
        try:
            # Calculate quantities
            quantities = self.calculate_cable_quantities()
            
            # Generate summary and detailed data
            summary_data = self.generate_summary_data(quantities)
            detailed_data = self.generate_detailed_data(quantities)
            
            # Create Excel writer
            writer = pd.ExcelWriter(filepath, engine='openpyxl')
            
            # Write summary data
            summary_data.to_excel(writer, sheet_name='Summary', index=False)
            
            # Write detailed data
            detailed_data.to_excel(writer, sheet_name='Block Details', index=False)
            
            # Get workbook
            workbook = writer.book
            
            # Format summary sheet
            self._format_excel_sheet(workbook['Summary'], summary_data)
            
            # Format detailed sheet
            self._format_excel_sheet(workbook['Block Details'], detailed_data)
            
            # Save the Excel file
            writer.close()
            
            return True
        except Exception as e:
            print(f"Error exporting BOM: {str(e)}")
            return False
    
    def _format_excel_sheet(self, worksheet, data: pd.DataFrame):
        """
        Format Excel worksheet
        
        Args:
            worksheet: openpyxl worksheet
            data: DataFrame with data
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
        
        # Format headers
        for col_num, column_title in enumerate(data.columns, 1):
            cell = worksheet.cell(row=1, column=col_num)
            cell.value = column_title
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = centered_alignment
            cell.border = border
        
        # Format data cells and apply auto-width
        for i, row in enumerate(worksheet.iter_rows(min_row=2, max_row=len(data)+1, min_col=1, max_col=len(data.columns)), 1):
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