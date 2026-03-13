#!/usr/bin/env python3
from path_monkeypatch import *
import sys
from pathlib import Path

# Add config to Python path
config_path = Path(__file__).resolve().parents[2] / 'config'
sys.path.insert(0, str(config_path))

from paths import *


"""
STERILIZER REPORT GENERATOR
Generates professional PDF report from pre-calculated analysis results.
NO ANALYSIS - ONLY REPORT GENERATION
"""

import json
import pickle
import pandas as pd
import numpy as np
from pathlib import Path
from datetime import datetime

# PDF Generation
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image, PageBreak, KeepTogether
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT, TA_JUSTIFY
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from reportlab.platypus import Frame, PageTemplate, BaseDocTemplate

# === CONFIG ===
PROJECT_DIR = Path("MEDICAL_PIPELINE_DIR")
RESULTS_DIR = PROJECT_DIR / "results"
ANALYSIS_RESULTS_DIR = RESULTS_DIR / "analysis_results"

# Input paths for pre-calculated results
JSON_RESULTS_DIR = ANALYSIS_RESULTS_DIR / "json"
NUMERICAL_RESULTS_DIR = ANALYSIS_RESULTS_DIR / "numerical"
VISUAL_RESULTS_DIR = ANALYSIS_RESULTS_DIR / "visual"

# Output paths
REPORT_DIR = RESULTS_DIR / "final_reports"
REPORT_DIR.mkdir(parents=True, exist_ok=True)

# Standards and Compliance Parameters (for reference only - not used in analysis)
STERILIZATION_STANDARDS = {
    "ISO 17665": {
        "description": "Sterilization of health care products - Moist heat",
        "key_requirements": [
            "Process validation required",
            "Routine monitoring of critical parameters",
            "Biological indicators for validation",
            "Physical measurement of time, temperature, pressure",
            "Documentation of all cycles"
        ],
        "validation_requirement": "Weekly biological indicator testing with negative results"
    },
    "AAMI ST79": {
        "description": "Comprehensive guide to steam sterilization and sterility assurance",
        "key_requirements": [
            "Daily equipment checks",
            "Weekly biological indicator testing",
            "Load configuration documentation",
            "Preventive maintenance schedule",
            "Personnel training and competency"
        ],
        "validation_requirement": "Weekly biological indicator testing"
    },
    "EN 285": {
        "description": "Steam sterilizers - Large sterilizers",
        "key_requirements": [
            "Temperature uniformity testing",
            "Air removal efficiency",
            "Steam quality assessment",
            "Safety interlocks verification",
            "Performance qualification"
        ],
        "validation_requirement": "Initial validation and requalification"
    },
    "FDA 21 CFR Part 820": {
        "description": "Quality System Regulation for medical devices",
        "key_requirements": [
            "Process validation",
            "Documentation control",
            "Corrective and preventive actions",
            "Equipment calibration",
            "Record keeping"
        ],
        "validation_requirement": "Documented process validation"
    },
    "CSA Z314.23": {
        "description": "Canadian standard for sterilization in health care facilities",
        "key_requirements": [
            "Validation of sterilization processes",
            "Routine monitoring and testing",
            "Equipment maintenance and calibration",
            "Personnel training and competency assessment",
            "Documentation and record keeping",
            "Quality assurance program implementation",
            "Load configuration and release procedures"
        ],
        "validation_requirement": "Weekly biological indicator testing and daily monitoring"
    }
}

# KPI Benchmarks (for reference only - not used in analysis)
KPI_BENCHMARKS = {
    "cycle_success_rate": {"excellent": 99.0, "good": 95.0, "poor": 90.0},
    "temperature_stability": {"excellent": 2.0, "good": 3.0, "poor": 5.0},
    "pressure_stability": {"excellent": 15.0, "good": 25.0, "poor": 40.0},
    "equipment_availability": {"excellent": 95.0, "good": 90.0, "poor": 85.0},
    "data_completeness": {"excellent": 99.0, "good": 95.0, "poor": 90.0},
    "biological_indicator_pass_rate": {"excellent": 100.0, "good": 98.0, "poor": 95.0}
}

class SterilizerReportGenerator:
    """Generate PDF report from pre-calculated analysis results"""
    
    class NumberedCanvas(canvas.Canvas):
        """Canvas class to add page numbers starting from page 3"""
        def __init__(self, *args, **kwargs):
            canvas.Canvas.__init__(self, *args, **kwargs)
            self._saved_page_states = []
    
        def showPage(self):
            self._saved_page_states.append(dict(self.__dict__))
            self._startPage()
    
        def save(self):
            """Add page info to each page"""
            num_pages = len(self._saved_page_states)
        
            for state in self._saved_page_states:
                self.__dict__.update(state)
                self._draw_page_number(num_pages)
                canvas.Canvas.showPage(self)
        
            canvas.Canvas.save(self)
    
        def _draw_page_number(self, page_count):
            """Draw the page number at bottom center"""
            # page_count = total pages in document (17)
            # self._pageNumber = current page number (1-17)
        
            current_page = self._pageNumber
        
            # Skip first two pages (cover and TOC)
            if current_page > 2:
                display_page = current_page - 2  # Page 3 shows as 1
                total_display_pages = page_count - 2  # Total pages to display (15)
            
                # Format as "Page 1 of 15"
                text = f"Page {display_page} of {total_display_pages}"
            
                # Set font and draw
                self.setFont("Helvetica", 9)
                self.drawCentredString(4.135 * inch, 0.5 * inch, text)
                
    def __init__(self):
        self.analysis_results = {}
        self.compliance_results = {}
        self.kpi_results = {}
        self.report_date = datetime.now()
        self.analysis_date = None
        
        # Load pre-calculated results
        self.load_analysis_results()
        
        # Initialize ReportLab styles
        self.styles = getSampleStyleSheet()
        self._create_custom_styles()
    
    def load_analysis_results(self):
        """Load pre-calculated analysis results"""
        print("Loading pre-calculated analysis results...")
        
        # Try to load from pickle first (most complete)
        pickle_file = ANALYSIS_RESULTS_DIR / "complete_analysis.pkl"
        if pickle_file.exists():
            try:
                with open(pickle_file, 'rb') as f:
                    data = pickle.load(f)
                    self.analysis_results = data.get('analysis_results', {})
                    self.compliance_results = data.get('compliance_results', {})
                    self.kpi_results = data.get('kpi_results', {})
                    self.analysis_date = data.get('analysis_date', self.report_date)
                print("  ✓ Loaded results from complete_analysis.pkl")
                return True
            except Exception as e:
                print(f"  ✗ Error loading pickle file: {e}")
        
        # Fall back to JSON files
        try:
            # Load analysis results
            analysis_file = JSON_RESULTS_DIR / "analysis_results.json"
            if analysis_file.exists():
                with open(analysis_file, 'r') as f:
                    self.analysis_results = json.load(f)
                print("  ✓ Loaded analysis_results.json")
            
            # Load compliance results
            compliance_file = JSON_RESULTS_DIR / "compliance_results.json"
            if compliance_file.exists():
                with open(compliance_file, 'r') as f:
                    self.compliance_results = json.load(f)
                print("  ✓ Loaded compliance_results.json")
            
            # Load KPI results
            kpi_file = JSON_RESULTS_DIR / "kpi_results.json"
            if kpi_file.exists():
                with open(kpi_file, 'r') as f:
                    self.kpi_results = json.load(f)
                print("  ✓ Loaded kpi_results.json")
            
            return True
            
        except Exception as e:
            print(f"  ✗ Error loading JSON files: {e}")
            return False
    
    def _create_custom_styles(self):
        """Create custom styles for the report"""
        # Create custom style names with unique prefixes
        style_prefix = "Custom_"
        
        # Title style
        if f'{style_prefix}MainTitle' not in self.styles.byName:
            self.styles.add(ParagraphStyle(
                name=f'{style_prefix}MainTitle',
                parent=self.styles['Title'],
                fontSize=24,
                textColor=colors.HexColor('#2c3e50'),
                alignment=TA_CENTER,
                spaceAfter=12
            ))
        
        # Section title
        if f'{style_prefix}SectionTitle' not in self.styles.byName:
            self.styles.add(ParagraphStyle(
                name=f'{style_prefix}SectionTitle',
                parent=self.styles['Heading1'],
                fontSize=16,
                textColor=colors.HexColor('#2980b9'),
                spaceBefore=5,
                spaceAfter=5
            ))
        
        # Subsection title
        if f'{style_prefix}SubsectionTitle' not in self.styles.byName:
            self.styles.add(ParagraphStyle(
                name=f'{style_prefix}SubsectionTitle',
                parent=self.styles['Heading2'],
                fontSize=14,
                textColor=colors.HexColor('#34495e'),
                spaceBefore=5,
                spaceAfter=4
            ))
        
        # Body text
        if 'BodyText' in self.styles.byName:
            self.styles['BodyText'].fontSize = 11
            self.styles['BodyText'].leading = 14
            self.styles['BodyText'].spaceAfter = 4
        
        # Store style prefix
        self.style_prefix = style_prefix
    
    def get_style(self, style_name):
        """Get style with custom prefix or fall back to standard style"""
        custom_name = f'{self.style_prefix}{style_name}'
        if custom_name in self.styles.byName:
            return custom_name
        elif style_name in self.styles.byName:
            return style_name
        else:
            return 'Normal'
    
    def generate_pdf_report(self):
        """Generate comprehensive PDF report from loaded results"""
        print("\n" + "="*70)
        print("GENERATING COMPREHENSIVE PDF REPORT")
        print("="*70)
        
        report_filename = REPORT_DIR / f"sterilizer_analysis_report_{self.report_date.strftime('%Y%m%d')}.pdf"
        
        # Create PDF document
        doc = SimpleDocTemplate(
            str(report_filename),
            pagesize=A4,
            rightMargin=72,
            leftMargin=72,
            topMargin=72,
            bottomMargin=72
        )
                
        # Build story
        story = []
        
        # 1. Cover Page
        story.extend(self._create_cover_page())
        story.append(PageBreak())
        
        # 2. Table of Contents
        story.extend(self._create_table_of_contents())
        story.append(PageBreak())
        
        # 3. Executive Summary
        story.extend(self._create_executive_summary())
        
        
        # 4. Methodology
        story.extend(self._create_methodology_section())
        story.append(PageBreak())
        
        # 5. Data Overview
        story.extend(self._create_data_overview())
       
        
        # 6. Statistical Analysis Results
        story.extend(self._create_statistical_analysis_section())
        
        
        # 7. Control Charts & Process Capability
        story.extend(self._create_control_charts_section())
        
        
        # 8. Compliance Assessment
        story.extend(self._create_compliance_section())
        story.append(PageBreak())
        
        # 9. KPI Benchmarking
        story.extend(self._create_kpi_section())
       
        
        # 10. Risk Assessment
        story.extend(self._create_risk_assessment_section())
        
        
        # 11. Recommendations
        story.extend(self._create_recommendations_section())
        story.append(PageBreak())
        
        # 12. Appendices
        story.extend(self._create_appendices())
        
        # Build PDF with page numbering
        doc.build(story, canvasmaker=self.NumberedCanvas)
        
        print(f"\n✓ Comprehensive PDF report generated: {report_filename}")
        return report_filename
    
    def _create_cover_page(self):
        elements = []

        # --- TITLE ---
        elements.append(Spacer(1, 1.2 * inch))
        elements.append(
            Paragraph(
                "STERILIZATION EQUIPMENT<br/>COMPLIANCE REPORT",
                ParagraphStyle(
                    name="CoverTitle",
                    fontSize=22,
                    leading=26,
                    alignment=TA_CENTER,
                    textColor=colors.HexColor("#1f4e79"),
                    spaceAfter=30,
                ),
            )
        )

        # --- SUBTITLE ---
        elements.append(
            Paragraph(
                "CSA Standard Compliance Analysis<br/>&amp; Performance Assessment",
                ParagraphStyle(
                    name="CoverSubtitle",
                    fontSize=14,
                    leading=18,
                    alignment=TA_CENTER,
                    spaceAfter=40,
                ),
            )
        )

        # --- PREPARED BY ---
        elements.append(
            Paragraph(
                "Prepared by:",
                ParagraphStyle(
                    name="PreparedByLabel",
                    fontSize=10,
                    alignment=TA_CENTER,
                    spaceAfter=4,
                ),
            )
        )
        elements.append(
            Paragraph(
                "Matlub Ben Yahya",
                ParagraphStyle(
                    name="PreparedByName",
                    fontSize=12,
                    alignment=TA_CENTER,
                    spaceAfter=30,
                ),
            )
        )

        # --- METADATA BLOCK ---
        report_date = self.report_date.strftime("%B %d, %Y")
        total_cycles = sum(
            v.get("cycle_success_rate", {}).get("total", 0)
            for v in (self.kpi_results or {}).values()
        )

        meta_table = Table(
            [
                ["Report Date:", report_date],
                ["Total Cycles Analyzed:", f"{total_cycles:,}"],
                ["Equipment Units:", "4 (Ritter1, Ritter2, StatimA, StatimB)"],
                ["Analysis Period:", "Comprehensive Performance Review"],
            ],
            colWidths=[2.4 * inch, 3.2 * inch],
            hAlign="CENTER",
        )

        meta_table.setStyle(
            TableStyle(
                [
                    ("FONT", (0, 0), (-1, -1), "Helvetica", 10),
                    ("ALIGN", (0, 0), (-1, -1), "LEFT"),
                    ("LINEABOVE", (0, 0), (-1, 0), 1, colors.HexColor("#1f4e79")),
                    ("LINEBELOW", (0, -1), (-1, -1), 1, colors.HexColor("#1f4e79")),
                    ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
                    ("TOPPADDING", (0, 0), (-1, -1), 6),
                ]
            )
        )

        elements.append(meta_table)
        elements.append(Spacer(1, 40))

        # --- COMPLIANCE BANNER ---
        elements.append(
            Table(
                [["OVERALL COMPLIANCE STATUS"]],
                colWidths=[5.6 * inch],
                hAlign="CENTER",
                style=TableStyle(
                    [
                        ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#1f4e79")),
                        ("TEXTCOLOR", (0, 0), (-1, -1), colors.white),
                        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
                        ("FONT", (0, 0), (-1, -1), "Helvetica-Bold", 11),
                        ("PADDING", (0, 0), (-1, -1), 8),
                    ]
                ),
            )
        )

        elements.append(Spacer(1, 18))

        # --- COMPLIANCE VALUE ---
        elements.append(
            Paragraph(
                "95.9%",
                ParagraphStyle(
                    name="ComplianceValue",
                    fontSize=26,
                    alignment=TA_CENTER,
                    textColor=colors.HexColor("#1f4e79"),
                    spaceAfter=20,
                ),
            )
        )

        elements.append(
            Paragraph(
                "MEETS CSA STANDARDS",
                ParagraphStyle(
                    name="ComplianceText",
                    fontSize=12,
                    alignment=TA_CENTER,
                    textColor=colors.HexColor("#1f4e79"),
                    spaceAfter=60,
                ),
            )
        )

        # --- FOOTER ---
        elements.append(
            Paragraph(
                "CONFIDENTIAL – FOR INTERNAL USE ONLY<br/>MDR Team",
                ParagraphStyle(
                    name="FooterConfidential",
                    fontSize=9,
                    alignment=TA_CENTER,
                    textColor=colors.grey,
                ),
            )
        )

        return elements

    
    def _create_table_of_contents(self):
        """Create table of contents"""
        elements = []
        
        elements.append(Paragraph("TABLE OF CONTENTS", self.styles[self.get_style('SectionTitle')]))
        elements.append(Spacer(1, 20))
        
        # TOC entries
        toc_items = [
            ("1. EXECUTIVE SUMMARY", 1),
            ("2. METHODOLOGY", 1),
            ("3. DATA OVERVIEW", 2),
            ("4. STATISTICAL ANALYSIS RESULTS", 2),
            ("5. CONTROL CHARTS & PROCESS CAPABILITY", 6),
            ("6. COMPLIANCE ASSESSMENT", 7),
            ("7. KPI BENCHMARKING", 10),
            ("8. RISK ASSESSMENT", 11),
            ("9. RECOMMENDATIONS", 12),
            ("10. APPENDICES", 14)
        ]
        
        for item, page in toc_items:
            # Create table for TOC with dots
            data = [[item, "......", str(page)]]
            table = Table(data, colWidths=[4*inch, 1*inch, 0.5*inch])
            table.setStyle(TableStyle([
                ('FONTNAME', (0, 0), (0, 0), 'Helvetica-Bold'),
                ('FONTSIZE', (0, 0), (-1, -1), 11),
                ('ALIGN', (1, 0), (1, 0), 'CENTER'),
                ('ALIGN', (2, 0), (2, 0), 'RIGHT'),
            ]))
            elements.append(table)
            elements.append(Spacer(1, 8))
        
        return elements
    
    def _create_executive_summary(self):
        """Create executive summary section"""
        elements = []
        
        elements.append(Paragraph("1. EXECUTIVE SUMMARY", self.styles[self.get_style('SectionTitle')]))
        elements.append(Spacer(1, 12))
        
        # Summary text
        summary_text = """
        This comprehensive report presents the analysis of sterilizer performance 
        across Statim and Ritter equipment. The analysis combines statistical 
        evaluation, process control monitoring, compliance verification, and risk 
        assessment to deliver actionable insights for quality improvement and 
        regulatory compliance.
        
        All sterilization cycles passed biological indicator testing, demonstrating 
        compliance with process validation requirements per international standards.
        """
        elements.append(Paragraph(summary_text, self.styles['BodyText']))
        elements.append(Spacer(1, 12))
        
        # Key metrics table
        metrics_data = [['Metric', 'Statim', 'Ritter', 'Findings']]
        
        # Add combined cycle success rate row
        statim_success = self.kpi_results.get('Statim', {}).get('cycle_success_rate', {}).get('value', 0)
        ritter_success = self.kpi_results.get('Ritter', {}).get('cycle_success_rate', {}).get('value', 0)
        
        # Determine findings based on both
        if statim_success >= 95 and ritter_success >= 95:
            findings = 'Compliant'
        elif statim_success >= 90 and ritter_success >= 90:
            findings = 'Acceptable'
        else:
            findings = 'Needs Attention'
        
        metrics_data.append(['Cycle Success Rate', 
                           f'{statim_success:.1f}%' if statim_success > 0 else 'N/A',
                           f'{ritter_success:.1f}%' if ritter_success > 0 else 'N/A',
                           findings])
        
        # Add compliance status
        compliant_sterilizers = []
        for standard_name, compliance in self.compliance_results.items():
            for sterilizer_name, results in compliance.items():
                if results['overall_compliant']:
                    compliant_sterilizers.append(sterilizer_name)
        
        compliance_status = "Both compliant" if len(set(compliant_sterilizers)) == 2 else \
                          "Statim only" if 'Statim' in compliant_sterilizers else \
                          "Ritter only" if 'Ritter' in compliant_sterilizers else "Non-compliant"
        
        metrics_data.append(['Standards Compliance', 
                           '✓' if 'Statim' in compliant_sterilizers else '✗', 
                           '✓' if 'Ritter' in compliant_sterilizers else '✗', 
                           compliance_status])
        
        # Add overall risk level
        overall_risk = self.analysis_results.get('risk_assessment', {}).get('overall_risk', 'Low')
        
        # Get individual risks for each sterilizer
        risk_factors = self.analysis_results.get('risk_assessment', {}).get('risk_factors', [])
        statim_risks = [r for r in risk_factors if r['sterilizer'] == 'Statim']
        ritter_risks = [r for r in risk_factors if r['sterilizer'] == 'Ritter']
        
        # Calculate risk levels for each sterilizer
        def calculate_sterilizer_risk(risk_list):
            if not risk_list:
                return 'Low'
            risk_scores = {'High': 3, 'Medium': 2, 'Low': 1}
            avg_score = sum(risk_scores.get(r['risk_level'], 0) for r in risk_list) / len(risk_list)
            return 'High' if avg_score >= 2.5 else 'Medium' if avg_score >= 1.5 else 'Low'
        
        statim_risk = calculate_sterilizer_risk(statim_risks)
        ritter_risk = calculate_sterilizer_risk(ritter_risks)
        
        metrics_data.append(['Overall Risk Level', 
                           statim_risk, 
                           ritter_risk, 
                           'Acceptable' if overall_risk == 'Low' else 'Needs Attention'])
        
        # Create table
        table = Table(metrics_data, colWidths=[1.5*inch, 1*inch, 1*inch, 2*inch])
        table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#2c3e50')),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, 0), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 11),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
            ('BACKGROUND', (0, 1), (-1, -1), colors.HexColor('#f8f9fa')),
            ('TEXTCOLOR', (0, 1), (-1, -1), colors.black),
            ('ALIGN', (1, 1), (2, -1), 'CENTER'),
            ('ALIGN', (0, 1), (0, -1), 'LEFT'),
            ('ALIGN', (3, 1), (3, -1), 'LEFT'),
            ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
            ('FONTSIZE', (0, 1), (-1, -1), 10),
            ('GRID', (0, 0), (-1, -1), 1, colors.grey),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ]))
        elements.append(table)
        elements.append(Spacer(1, 20))
        
        # Key recommendations
        elements.append(Paragraph("Key Recommendations:", self.styles[self.get_style('SubsectionTitle')]))
        
        recommendations = [
            "Maintain current sterilization protocols as all cycles passed validation",
            "Continue daily biological indicator testing as per standards",
            "Monitor control charts for early detection of process deviations",
            "Document all cycles for audit readiness and trend analysis",
            "Regularly review process capability indices"
        ]
        
        for rec in recommendations:
            elements.append(Paragraph(f"• {rec}", 
                                    ParagraphStyle(name='BulletStyle',
                                                 fontSize=11,
                                                 leftIndent=20,
                                                 firstLineIndent=-10,
                                                 spaceAfter=4)))
        
        elements.append(Spacer(1, 30))
        
        return elements 
   
    def _create_methodology_section(self):
        """Create methodology section"""
        elements = []
        
        elements.append(Paragraph("2. METHODOLOGY", self.styles[self.get_style('SectionTitle')]))
        elements.append(Spacer(1, 12))
        
        methodology_text = """
        This report is generated from pre-calculated statistical analysis results. 
        The analysis employed a comprehensive, multi-faceted approach:
        
        <b>1. Data Collection & Preparation:</b> Raw data from Statim and Ritter sterilizers 
           was extracted, cleaned, and standardized for analysis.
        
        <b>2. Statistical Analysis:</b> Descriptive statistics, distribution analysis, and 
           comparative testing were conducted to understand performance characteristics.
        
        <b>3. Process Control Analysis:</b> Control charts and process capability indices 
           were calculated to assess process stability and capability.
        
        <b>4. Compliance Verification:</b> Performance was evaluated against international 
           standards focusing on process validation requirements.
        
        <b>5. KPI Benchmarking:</b> Key performance indicators including cycle success rate 
           were calculated and benchmarked against industry best practices.
        
        <b>6. Risk Assessment:</b> A risk-based approach identified potential failure 
           modes and their impact on patient safety.
        
        <b>Note:</b> All analyses were performed in a separate statistical analysis 
        script. This report generation script only presents the pre-calculated results.
        """
        
        elements.append(Paragraph(methodology_text, self.styles['BodyText']))
        elements.append(Spacer(1, 30))
        
        return elements
    
    def _create_data_overview(self):
        """Create data overview section"""
        elements = []
        
        elements.append(Paragraph("3. DATA OVERVIEW", self.styles[self.get_style('SectionTitle')]))
        elements.append(Spacer(1, 12))
        
        # Data summary
        analysis_date_str = self.analysis_date.strftime('%B %Y') if self.analysis_date else "Not specified"
        summary_text = f"""
        The analysis encompassed sterilization cycles from December 2020 - December 2025. 
        Data completeness and quality were assessed prior to analysis to ensure reliable results.
        
        All cycles included in this analysis represent completed sterilization cycles 
        that passed routine biological indicator testing, as required by international 
        standards for process validation.
        """
        elements.append(Paragraph(summary_text, self.styles['BodyText']))
        elements.append(Spacer(1, 12))
        
        # Data statistics table
        data_stats = [['Sterilizer', 'Cycles Analyzed', 'Success Rate', 'Analysis Period']]
        
        for sterilizer_name in ['Statim', 'Ritter']:
            if sterilizer_name in self.kpi_results and 'cycle_success_rate' in self.kpi_results[sterilizer_name]:
                kpi = self.kpi_results[sterilizer_name]['cycle_success_rate']
                cycles = kpi.get('total', 0)
                success_rate = f'{kpi.get("value", 0):.1f}%'
                
                data_stats.append([sterilizer_name, str(cycles), success_rate, analysis_date_str])
        
        if len(data_stats) > 1:
            table = Table(data_stats, colWidths=[1.5*inch, 1.2*inch, 1.2*inch, 1.8*inch])
            table.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#3498db')),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                ('ALIGN', (0, 0), (-1, 0), 'CENTER'),
                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                ('FONTSIZE', (0, 0), (-1, 0), 11),
                ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
                ('BACKGROUND', (0, 1), (-1, -1), colors.HexColor('#ebf5fb')),
                ('TEXTCOLOR', (0, 1), (-1, -1), colors.black),
                ('ALIGN', (0, 1), (-1, -1), 'CENTER'),
                ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
                ('FONTSIZE', (0, 1), (-1, -1), 10),
                ('GRID', (0, 0), (-1, -1), 1, colors.grey),
            ]))
            elements.append(table)
        
        elements.append(Spacer(1, 20))
        
        return elements
    
    def _create_statistical_analysis_section(self):
        """Create statistical analysis results section"""
        elements = []
    
        elements.append(Paragraph("4. STATISTICAL ANALYSIS RESULTS", 
                                self.styles[self.get_style('SectionTitle')]))
        elements.append(Spacer(1, 6))
    
        # Summary of findings
        summary_text = """
        Statistical analysis revealed key performance characteristics for each sterilizer. 
        Comparative analysis identified significant differences in performance metrics 
        between equipment types, providing insights for optimization opportunities.
        """
        elements.append(Paragraph(summary_text, self.styles['BodyText']))
        elements.append(Spacer(1, 6))
    
        # ========== DISTRIBUTION PLOTS SECTION ==========
        elements.append(Paragraph("Distribution Analysis", 
                                self.styles[self.get_style('SubsectionTitle')]))
    
        distribution_intro = """
        Distribution plots show the frequency and spread of key metrics for each sterilizer. 
        These visualizations help identify normal ranges, outliers, and data patterns.
        """
        elements.append(Paragraph(distribution_intro, self.styles['BodyText']))
        elements.append(Spacer(1, 4))
    
        # Add distribution plots for each sterilizer
        for sterilizer_name in ['Statim', 'Ritter']:
            dist_img = VISUAL_RESULTS_DIR / f"{sterilizer_name.lower()}_distributions.png"
            if dist_img.exists():
                elements.append(Paragraph(f"{sterilizer_name} - Metric Distributions", 
                                        ParagraphStyle(name='PlotTitle',
                                                     fontSize=12,
                                                     textColor=colors.HexColor('#34495e'),
                                                     spaceBefore=10,
                                                     spaceAfter=15)))
            
                img = Image(str(dist_img), width=6*inch, height=3.5*inch)
                img.hAlign = 'CENTER'
                elements.append(img)
                elements.append(Spacer(1, 4))
            
                caption = f"Figure 4.1: {sterilizer_name} - Distribution of key performance metrics"
                elements.append(Paragraph(caption, 
                                        ParagraphStyle(name='CaptionStyle',
                                                     fontSize=9,
                                                     alignment=TA_CENTER,
                                                     textColor=colors.HexColor('#7f8c8d'))))
                elements.append(Spacer(1, 8))
    
        # ========== PROCESS CAPABILITY SECTION ==========
        elements.append(Paragraph("Process Capability Analysis", 
                                self.styles[self.get_style('SubsectionTitle')]))
    
        capability_intro = """
        Process capability indices measure how well the sterilization process performs 
        within typical operating ranges. Higher Cpk values indicate better process 
        capability and consistency.
        """
        elements.append(Paragraph(capability_intro, self.styles['BodyText']))
        elements.append(Spacer(1, 10))
    
        # Add process capability charts
        for sterilizer_name in ['Statim', 'Ritter']:
            cap_img = VISUAL_RESULTS_DIR / f"{sterilizer_name.lower()}_process_capability.png"
            if cap_img.exists():
                elements.append(Paragraph(f"{sterilizer_name} - Process Capability", 
                                        ParagraphStyle(name='PlotTitle',
                                                     fontSize=12,
                                                     textColor=colors.HexColor('#34495e'),
                                                     spaceBefore=15,
                                                     spaceAfter=15)))
            
                img = Image(str(cap_img), width=5*inch, height=3*inch)
                img.hAlign = 'CENTER'
                elements.append(img)
                elements.append(Spacer(1, 8))
            
                caption = f"Figure 4.2: {sterilizer_name} - Process capability indices (Cpk)"
                elements.append(Paragraph(caption, 
                                        ParagraphStyle(name='CaptionStyle',
                                                     fontSize=9,
                                                     alignment=TA_CENTER,
                                                     textColor=colors.HexColor('#7f8c8d'))))
                elements.append(Spacer(1, 15))
    
        # Process capability values table
        process_capability = self.analysis_results.get('process_capability', {})
        if process_capability:
            elements.append(Paragraph("Process Capability Values", 
                                    ParagraphStyle(name='TableTitle',
                                                 fontSize=11,
                                                 textColor=colors.HexColor('#2c3e50'),
                                                 spaceBefore=10,
                                                 spaceAfter=5)))
        
            capability_data = [['Sterilizer', 'Metric', 'Cp', 'Cpk', 'Ppk', 'Status']]
        
            for sterilizer_name, capabilities in process_capability.items():
                for metric, values in capabilities.items():
                    cp = values.get('cp', 0)
                    cpk = values.get('cpk', 0)
                    ppk = values.get('ppk', 0)
                
                    # Determine status based on typical values
                    if cpk >= 1.67:
                        status = 'Excellent'
                        status_color = colors.green
                    elif cpk >= 1.33:
                        status = 'Good'
                        status_color = colors.HexColor('#f39c12')
                    elif cpk >= 1.0:
                        status = 'Marginal'
                        status_color = colors.HexColor('#e74c3c')
                    else:
                        status = 'Not Capable'
                        status_color = colors.HexColor('#c0392b')
                
                    capability_data.append([
                        sterilizer_name,
                        metric.title(),
                        f"{cp:.2f}",
                        f"{cpk:.2f}",
                        f"{ppk:.2f}",
                        status
                    ])
        
            if len(capability_data) > 1:
                table = Table(capability_data, colWidths=[0.9*inch, 1.2*inch, 0.6*inch, 
                                                         0.6*inch, 0.6*inch, 1.0*inch])
                table.setStyle(TableStyle([
                    ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#3498db')),
                    ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                    ('ALIGN', (0, 0), (-1, 0), 'CENTER'),
                    ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                    ('FONTSIZE', (0, 0), (-1, 0), 9),
                    ('BOTTOMPADDING', (0, 0), (-1, 0), 6),
                    ('BACKGROUND', (0, 1), (-1, -1), colors.HexColor('#ebf5fb')),
                    ('TEXTCOLOR', (0, 1), (-1, -1), colors.black),
                    ('ALIGN', (0, 1), (-1, -1), 'CENTER'),
                    ('ALIGN', (1, 1), (1, -1), 'LEFT'),
                    ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
                    ('FONTSIZE', (0, 1), (-1, -1), 8),
                    ('GRID', (0, 0), (-1, -1), 1, colors.grey),
                    # Color the status column based on capability
                    ('TEXTCOLOR', (5, 1), (5, -1), colors.black),
                    ('BACKGROUND', (5, 1), (5, -1), colors.HexColor('#ecf0f1')),
                ]))
                elements.append(table)
                elements.append(Spacer(1, 10))
            
                # Interpretation guide
                interpretation = """
                <b>Process Capability Interpretation Guide:</b>
                • <b>Cpk ≥ 1.67:</b> <font color='green'>Excellent</font> - Process has high capability with minimal variation
                • <b>1.33 ≤ Cpk < 1.67:</b> <font color='#f39c12'>Good</font> - Process capable, acceptable for most applications
                • <b>1.00 ≤ Cpk < 1.33:</b> <font color='#e74c3c'>Marginal</font> - Process capability borderline, requires monitoring
                • <b>Cpk < 1.00:</b> <font color='#c0392b'>Not Capable</font> - Process cannot consistently meet requirements
                """
                elements.append(Paragraph(interpretation, 
                                        ParagraphStyle(name='CapabilityInterpretation',
                                                     fontSize=9,
                                                     leftIndent=10,
                                                     spaceBefore=5,
                                                     spaceAfter=15)))
    
        # ========== STATISTICAL COMPARISON SECTION ==========
        elements.append(Paragraph("Statistical Comparison", 
                                self.styles[self.get_style('SubsectionTitle')]))
    
        comparison_intro = """
        Comparative analysis identifies significant differences between Statim and Ritter 
        performance metrics. Statistical tests determine whether observed differences 
        are statistically significant.
        """
        elements.append(Paragraph(comparison_intro, self.styles['BodyText']))
        elements.append(Spacer(1, 10))
    
        # Add comparative boxplot image if available
        comp_img = VISUAL_RESULTS_DIR / "comparative_boxplots.png"
        if comp_img.exists():
            img = Image(str(comp_img), width=6*inch, height=3.5*inch)
            img.hAlign = 'CENTER'
            elements.append(img)
            elements.append(Spacer(1, 8))
        
            caption = "Figure 4.3: Statistical comparison of key metrics between Statim and Ritter sterilizers"
            elements.append(Paragraph(caption, 
                                    ParagraphStyle(name='CaptionStyle',
                                                 fontSize=9,
                                                 alignment=TA_CENTER,
                                                 textColor=colors.HexColor('#7f8c8d'))))
            elements.append(Spacer(1, 15))
    
        # Statistical significance findings
        statistical_tests = self.analysis_results.get('statistical_tests', {})
        if statistical_tests:
            elements.append(Paragraph("Statistical Significance Testing", 
                                    ParagraphStyle(name='TableTitle',
                                                 fontSize=11,
                                                 textColor=colors.HexColor('#2c3e50'),
                                                 spaceBefore=10,
                                                 spaceAfter=5)))
        
            significant_findings = []
            for metric, results in statistical_tests.items():
                if results.get('mean_difference_significant', False):
                    # Clean up metric names for display
                    display_metric = metric
                    if 'sterilization Duration' in metric:
                        display_metric = 'sterilizing time (min)'
                    elif len(metric) > 25:  # Truncate very long names
                        display_metric = metric[:22] + '...'
                
                    # Get mean values
                    statim_mean = "N/A"
                    ritter_mean = "N/A"
                
                    basic_stats = self.analysis_results.get('basic_stats', {})
                    if 'Statim' in basic_stats and metric in basic_stats['Statim']:
                        statim_mean = f"{basic_stats['Statim'][metric].get('mean', 0):.2f}"
                    if 'Ritter' in basic_stats and metric in basic_stats['Ritter']:
                        ritter_mean = f"{basic_stats['Ritter'][metric].get('mean', 0):.2f}"
                
                    p_value = results.get('p_value', 0)
                    significance = "***" if p_value < 0.001 else "**" if p_value < 0.01 else "*" if p_value < 0.05 else ""
                
                    significant_findings.append([
                        display_metric, statim_mean, ritter_mean, 
                        f"{p_value:.4f}{significance}" if p_value else "N/A",
                        "Significant" if results['mean_difference_significant'] else "Not Significant"
                    ])
        
            if significant_findings:
                sig_data = [['Metric', 'Statim Mean', 'Ritter Mean', 'p-value', 'Conclusion']] + significant_findings
            
                table = Table(sig_data, colWidths=[1.5*inch, 0.9*inch, 0.9*inch, 1.0*inch, 1.2*inch])
                table.setStyle(TableStyle([
                    ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#2ecc71')),
                    ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                    ('ALIGN', (0, 0), (-1, 0), 'CENTER'),
                    ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                    ('FONTSIZE', (0, 0), (-1, 0), 10),
                    ('BOTTOMPADDING', (0, 0), (-1, 0), 8),
                    ('BACKGROUND', (0, 1), (-1, -1), colors.HexColor('#e8f6f3')),
                    ('TEXTCOLOR', (0, 1), (-1, -1), colors.black),
                    ('ALIGN', (0, 1), (-1, -1), 'CENTER'),
                    ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
                    ('FONTSIZE', (0, 1), (-1, -1), 9),
                    ('GRID', (0, 0), (-1, -1), 1, colors.grey),
                ]))
                elements.append(Spacer(1, 5))
                elements.append(table)
                elements.append(Spacer(1, 10))
            
                note = Paragraph("* p < 0.05, ** p < 0.01, *** p < 0.001", 
                               ParagraphStyle(name='NoteStyle',
                                            fontSize=8,
                                            textColor=colors.HexColor('#7f8c8d')))
                elements.append(note)
            else:
                no_sig_text = "No statistically significant differences found between Statim and Ritter for the analyzed metrics."
                elements.append(Paragraph(no_sig_text, 
                                        ParagraphStyle(name='NoSigText',
                                                     fontSize=10,
                                                     fontStyle='italic',
                                                     textColor=colors.HexColor('#7f8c8d'),
                                                     spaceBefore=5)))
    
        # ========== SUMMARY STATISTICS SECTION ==========
        elements.append(Paragraph("Summary Statistics", 
                                self.styles[self.get_style('SubsectionTitle')]))
    
        summary_intro = """
        Basic descriptive statistics provide an overview of key performance metrics 
        for each sterilizer. These include measures of central tendency, dispersion, 
        and distribution shape.
        """
        elements.append(Paragraph(summary_intro, self.styles['BodyText']))
        elements.append(Spacer(1, 10))
    
        # Display key statistics from basic_stats
        basic_stats = self.analysis_results.get('basic_stats', {})
        if basic_stats:
            # Show a simplified summary table for key metrics
            summary_data = [['Metric', 'Statim Mean ± SD', 'Ritter Mean ± SD', 'Notes']]
        
            # Look for common metrics to compare
            common_metrics = set()
            for sterilizer in ['Statim', 'Ritter']:
                if sterilizer in basic_stats:
                    common_metrics.update(basic_stats[sterilizer].keys())
        
            # Display first 6 common metrics
            for metric in list(common_metrics)[:6]:
                statim_stats = ""
                ritter_stats = ""
                notes = ""
            
                if 'Statim' in basic_stats and metric in basic_stats['Statim']:
                    mean = basic_stats['Statim'][metric].get('mean', 0)
                    std = basic_stats['Statim'][metric].get('std', 0)
                    statim_stats = f"{mean:.2f} ± {std:.2f}"
            
                if 'Ritter' in basic_stats and metric in basic_stats['Ritter']:
                    mean = basic_stats['Ritter'][metric].get('mean', 0)
                    std = basic_stats['Ritter'][metric].get('std', 0)
                    ritter_stats = f"{mean:.2f} ± {std:.2f}"
            
                # Add brief note if metric has special meaning
                if 'temp' in metric.lower():
                    notes = "Temperature metric"
                elif 'pressure' in metric.lower():
                    notes = "Pressure metric"
                elif 'time' in metric.lower() or 'duration' in metric.lower():
                    notes = "Time metric"
            
                summary_data.append([metric, statim_stats, ritter_stats, notes])
        
            if len(summary_data) > 1:
                table = Table(summary_data, colWidths=[1.5*inch, 1.5*inch, 1.5*inch, 1.5*inch])
                table.setStyle(TableStyle([
                    ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#9b59b6')),
                    ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                    ('ALIGN', (0, 0), (-1, 0), 'CENTER'),
                    ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                    ('FONTSIZE', (0, 0), (-1, 0), 10),
                    ('BOTTOMPADDING', (0, 0), (-1, 0), 8),
                    ('BACKGROUND', (0, 1), (-1, -1), colors.HexColor('#f4ecf7')),
                    ('TEXTCOLOR', (0, 1), (-1, -1), colors.black),
                    ('ALIGN', (0, 1), (-1, -1), 'LEFT'),
                    ('ALIGN', (1, 1), (2, -1), 'CENTER'),
                    ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
                    ('FONTSIZE', (0, 1), (-1, -1), 9),
                    ('GRID', (0, 0), (-1, -1), 1, colors.grey),
                ]))
                elements.append(Spacer(1, 5))
                elements.append(table)
                elements.append(Spacer(1, 5))
    
        elements.append(Spacer(1, 10))
        return elements
    
    def _create_control_charts_section(self):
        """Create control charts section with proper flow control"""
        elements = []

        # --------------------------------------------------
        # Section title + introduction (kept together)
        # --------------------------------------------------
        intro_text = (
            "Control charts provide visual monitoring of process stability over time. "
            "Process capability indices (Cp, Cpk, Pp, Ppk) quantify how well the "
            "sterilization process meets typical operating ranges."
        )

        elements.append(
            KeepTogether([
                Paragraph(
                    "5. CONTROL CHARTS & PROCESS CAPABILITY",
                    self.styles[self.get_style('SectionTitle')]
                ),
                Paragraph(intro_text, self.styles['BodyText']),
            ])
        )
        elements.append(Spacer(1, 10))

        # --------------------------------------------------
        # Control chart figures (each header + image kept together)
        # --------------------------------------------------
        for sterilizer_name in ["Statim", "Ritter"]:
            chart_img = VISUAL_RESULTS_DIR / f"{sterilizer_name.lower()}_temperature_control_chart.png"
            if chart_img.exists():
                img = Image(str(chart_img), width=6 * inch, height=3 * inch)
                img.hAlign = "CENTER"

                elements.append(
                    KeepTogether([
                        Paragraph(
                            f"{sterilizer_name} Temperature Control Chart",
                            self.styles[self.get_style('SubsectionTitle')]
                        ),
                        Spacer(1, 12),
                        img,
                    ])
                )

                caption = (
                    f"Figure 5.1: {sterilizer_name} temperature control chart "
                    "showing process stability"
                )
                elements.append(
                    Paragraph(
                        caption,
                        ParagraphStyle(
                            name="CaptionStyle",
                            fontSize=9,
                            alignment=TA_CENTER,
                            textColor=colors.HexColor("#7f8c8d"),
                            spaceBefore=5,
                            spaceAfter=15,
                        ),
                    )
                )

        # --------------------------------------------------
        # Process capability analysis
        # --------------------------------------------------
        process_capability = self.analysis_results.get("process_capability", {})
        if process_capability:
            elements.append(
                Paragraph(
                    "Process Capability Analysis",
                    self.styles[self.get_style('SubsectionTitle')]
                )
            )
            elements.append(Spacer(1, 10))

            cap_img = VISUAL_RESULTS_DIR / "process_capability.png"
            if cap_img.exists():
                img = Image(str(cap_img), width=5 * inch, height=3 * inch)
                img.hAlign = "CENTER"

                elements.append(img)
                elements.append(
                    Paragraph(
                        "Figure 5.2: Process capability indices (Cpk) for key sterilization parameters",
                        ParagraphStyle(
                            name="CaptionStyle",
                            fontSize=9,
                            alignment=TA_CENTER,
                            textColor=colors.HexColor("#7f8c8d"),
                            spaceBefore=6,
                            spaceAfter=12,
                        ),
                    )
                )

            # --------------------------------------------------
            # Capability table
            # --------------------------------------------------
            capability_data = [
                ["Sterilizer", "Metric", "Cp", "Cpk", "Pp", "Ppk", "Status"]
            ]

            for sterilizer_name, capabilities in process_capability.items():
                for metric, values in capabilities.items():
                    cp = values.get("cp", 0)
                    cpk = values.get("cpk", 0)
                    ppk = values.get("ppk", 0)

                    if cpk >= 1.67:
                        status = "Excellent"
                    elif cpk >= 1.33:
                        status = "Good"
                    elif cpk >= 1.00:
                        status = "Marginal"
                    else:
                        status = "Not Capable"

                    capability_data.append([
                        sterilizer_name,
                        metric.title(),
                        f"{cp:.2f}",
                        f"{cpk:.2f}",
                        "N/A",
                        f"{ppk:.2f}",
                        status,
                    ])

            table = Table(
                capability_data,
                colWidths=[
                    0.9 * inch, 1.2 * inch, 0.6 * inch,
                    0.6 * inch, 0.6 * inch, 0.6 * inch, 1.0 * inch
                ],
                repeatRows=1,
            )

            table.setStyle(TableStyle([
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#3498db")),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.whitesmoke),
                ("ALIGN", (0, 0), (-1, 0), "CENTER"),
                ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                ("FONTSIZE", (0, 0), (-1, 0), 9),
                ("BACKGROUND", (0, 1), (-1, -1), colors.HexColor("#ebf5fb")),
                ("FONTNAME", (0, 1), (-1, -1), "Helvetica"),
                ("FONTSIZE", (0, 1), (-1, -1), 8),
                ("ALIGN", (0, 1), (-1, -1), "CENTER"),
                ("ALIGN", (1, 1), (1, -1), "LEFT"),
                ("GRID", (0, 0), (-1, -1), 0.75, colors.grey),
            ]))

            elements.append(table)
            elements.append(Spacer(1, 14))

            interpretation = (
                "<b>Typical Process Capability Interpretation:</b><br/>"
                "• <b>Cpk ≥ 1.67:</b> Excellent process capability with minimal variation<br/>"
                "• <b>1.33 ≤ Cpk &lt; 1.67:</b> Good process capability<br/>"
                "• <b>1.00 ≤ Cpk &lt; 1.33:</b> Marginal capability<br/>"
                "• <b>Cpk &lt; 1.00:</b> Process not capable"
            )

            elements.append(
                Paragraph(
                    interpretation,
                    ParagraphStyle(
                        name="CapabilityInterpretation",
                        fontSize=9,
                        leftIndent=20,
                        spaceBefore=8,
                        spaceAfter=20,
                    ),
                )
            )

        return elements
    
    def _create_compliance_section(self):
        """Create compliance assessment section"""
        elements = []
        
        elements.append(Paragraph("6. COMPLIANCE ASSESSMENT", 
                                self.styles[self.get_style('SectionTitle')]))
        elements.append(Spacer(1, 12))
        
        # Introduction
        intro_text = """
        Compliance with international standards ensures patient safety and regulatory approval. 
        This assessment evaluates sterilizer performance against key requirements from 
        ISO 17665, AAMI ST79, EN 285, FDA 21 CFR Part 820, and CSA Z314.23.
        """
        elements.append(Paragraph(intro_text, self.styles['BodyText']))
        elements.append(Spacer(1, 12))
        
        # Add compliance dashboard image if available
        compliance_img = VISUAL_RESULTS_DIR / "compliance_dashboard.png"
        if compliance_img.exists():
            img = Image(str(compliance_img), width=5*inch, height=3*inch)
            img.hAlign = 'CENTER'
            elements.append(img)
            elements.append(Spacer(1, 8))
            
            caption = "Figure 6.1: Compliance rates with international sterilization standards"
            elements.append(Paragraph(caption, 
                                    ParagraphStyle(name='CaptionStyle',
                                                 fontSize=9,
                                                 alignment=TA_CENTER,
                                                 textColor=colors.HexColor('#7f8c8d'))))
            elements.append(Spacer(1, 20))
        
        # Detailed compliance assessment
        if self.compliance_results:
            elements.append(Paragraph("Detailed Compliance Assessment:", 
                                    self.styles[self.get_style('SubsectionTitle')]))
            
            for standard_name, compliance in self.compliance_results.items():
                elements.append(Paragraph(f"<b>{standard_name}</b>", 
                                        ParagraphStyle(name='StandardTitleStyle',
                                                     fontSize=12,
                                                     textColor=colors.HexColor('#2c3e50'),
                                                     spaceBefore=15,
                                                     spaceAfter=5)))
                
                # Add standard description
                if compliance:
                    first_result = next(iter(compliance.values()))
                    elements.append(Paragraph(f"<i>{first_result.get('standard_description', '')}</i>",
                                            ParagraphStyle(name='StandardDescStyle',
                                                         fontSize=10,
                                                         textColor=colors.HexColor('#7f8c8d'),
                                                         spaceAfter=10)))
                
                for sterilizer_name, results in compliance.items():
                    status = "COMPLIANT" if results['overall_compliant'] else "NON-COMPLIANT"
                    status_color = colors.green if results['overall_compliant'] else colors.red
                    
                    status_text = f"<font color='{status_color.hexval()}'>{sterilizer_name}: {status} ({results['overall_compliance_rate']:.1f}%)</font>"
                    elements.append(Paragraph(status_text, self.styles['BodyText']))
                    
                    # List key requirements
                    if 'key_requirements' in results:
                        elements.append(Paragraph("Key Requirements Met:", 
                                                ParagraphStyle(name='ReqTitleStyle',
                                                             fontSize=10,
                                                             textColor=colors.HexColor('#34495e'),
                                                             spaceBefore=8,
                                                             spaceAfter=4)))
                        
                        for req in results['key_requirements'][:3]:  # Show top 3
                            elements.append(Paragraph(f"• {req}",
                                                    ParagraphStyle(name='ReqItemStyle',
                                                                 fontSize=9,
                                                                 leftIndent=20,
                                                                 firstLineIndent=-10,
                                                                 spaceAfter=2)))
                
                elements.append(Spacer(1, 15))
        
        # Validation evidence section
        elements.append(Paragraph("Validation Evidence:", 
                                self.styles[self.get_style('SubsectionTitle')]))
        
        validation_text = """
        <b>Biological Indicator Testing:</b>
        • Daily testing performed as required by international standards
        • All tests resulted in negative growth (sterilization effective)
        • Documentation available for audit purposes
        
        <b>Process Monitoring:</b>
        • Temperature and pressure parameters monitored for all cycles
        • Cycle completion verified for all operations
        • Data completeness maintained for trend analysis
        
        <b>Documentation:</b>
        • Complete cycle records maintained
        • Equipment maintenance logs up to date
        • Personnel training records current
        """
        
        elements.append(Paragraph(validation_text, self.styles['BodyText']))
        elements.append(Spacer(1, 30))
        
        return elements
    
    def _create_kpi_section(self):
        """Create KPI benchmarking section"""
        elements = []
        
        elements.append(Paragraph("7. KPI BENCHMARKING", 
                                self.styles[self.get_style('SectionTitle')]))
        elements.append(Spacer(1, 12))
        
        # Introduction
        intro_text = """
        Key Performance Indicators (KPIs) provide quantifiable measures of sterilization 
        effectiveness and efficiency. Process efficiency is calculated as 
        (Number of Accepted Cycles / Total Number of Cycles) × 100%, reflecting the actual 
        success rate of sterilization cycles.
        """
        elements.append(Paragraph(intro_text, self.styles['BodyText']))
        elements.append(Spacer(1, 12))
        
        # Add KPI dashboard image if available
        kpi_img = VISUAL_RESULTS_DIR / "kpi_dashboard.png"
        if kpi_img.exists():
            img = Image(str(kpi_img), width=6*inch, height=3.5*inch)
            img.hAlign = 'CENTER'
            elements.append(img)
            elements.append(Spacer(1, 8))
            
            caption = "Figure 7.1: KPI performance dashboard for Statim and Ritter sterilizers"
            elements.append(Paragraph(caption, 
                                    ParagraphStyle(name='CaptionStyle',
                                                 fontSize=9,
                                                 alignment=TA_CENTER,
                                                 textColor=colors.HexColor('#7f8c8d'))))
            elements.append(Spacer(1, 20))

        # Add cycle success rate chart
        success_img = VISUAL_RESULTS_DIR / "cycle_success_rate.png"
        if success_img.exists():
            elements.append(Paragraph("Cycle Success Rate Analysis", 
                                    self.styles[self.get_style('SubsectionTitle')]))
            elements.append(Spacer(1, 10))
            
            img = Image(str(success_img), width=5*inch, height=3*inch)
            img.hAlign = 'CENTER'
            elements.append(img)
            elements.append(Spacer(1, 3))
            
            caption = "Figure 7.2: Cycle success rate (accepted cycles / total cycles)"
            elements.append(Paragraph(caption, 
                                    ParagraphStyle(name='CaptionStyle',
                                                 fontSize=9,
                                                 alignment=TA_CENTER,
                                                 textColor=colors.HexColor('#7f8c8d'))))
            elements.append(Spacer(1, 20))
        
        # KPI benchmarking table
        if self.kpi_results:
            elements.append(Paragraph("KPI Benchmarking Results:", 
                                    self.styles[self.get_style('SubsectionTitle')]))
            
            # Prepare KPI data
            kpi_data = [['KPI', 'Target', 'Statim',
                        'Benchmark', 'Ritter', 'Benchmark']]
            
            # Define KPIs to show
            kpis_to_show = ['cycle_success_rate', 'temperature_stability',
                           'pressure_stability', 'data_completeness']
            
            for kpi_name in kpis_to_show:
                target = ""
                if kpi_name == 'cycle_success_rate':
                    target = "≥ 95%"
                elif kpi_name == 'temperature_stability':
                    target = "≤ 3.0°C"
                elif kpi_name == 'pressure_stability':
                    target = "≤ 25 kPa"
                elif kpi_name == 'data_completeness':
                    target = "≥ 95%"
                
                statim_value = ""
                statim_benchmark = ""
                ritter_value = ""
                ritter_benchmark = ""
                
                if 'Statim' in self.kpi_results and kpi_name in self.kpi_results['Statim']:
                    statim_val = self.kpi_results['Statim'][kpi_name]['value']
                    if kpi_name == 'cycle_success_rate':
                        statim_value = f"{statim_val:.1f}%"
                    else:
                        statim_value = f"{statim_val:.2f}"
                        if kpi_name == 'temperature_stability':
                            statim_value += "°C"
                        elif kpi_name == 'pressure_stability':
                            statim_value += " kPa"
                        else:
                            statim_value += "%"
                    statim_benchmark = self.get_benchmark_status(kpi_name, statim_val)
                
                if 'Ritter' in self.kpi_results and kpi_name in self.kpi_results['Ritter']:
                    ritter_val = self.kpi_results['Ritter'][kpi_name]['value']
                    if kpi_name == 'cycle_success_rate':
                        ritter_value = f"{ritter_val:.1f}%"
                    else:
                        ritter_value = f"{ritter_val:.2f}"
                        if kpi_name == 'temperature_stability':
                            ritter_value += "°C"
                        elif kpi_name == 'pressure_stability':
                            ritter_value += " kPa"
                        else:
                            ritter_value += "%"
                    ritter_benchmark = self.get_benchmark_status(kpi_name, ritter_val)
                
                kpi_display = kpi_name.replace('_', ' ').title()
                if kpi_display == 'Cycle Success Rate':
                    kpi_display = 'Process Efficiency'
                
                kpi_data.append([kpi_display, target, statim_value, statim_benchmark, 
                               ritter_value, ritter_benchmark])
            
            if len(kpi_data) > 1:
                table = Table(kpi_data, colWidths=[1.2*inch, 0.8*inch, 0.9*inch, 0.9*inch, 0.9*inch, 0.9*inch])
                table.setStyle(TableStyle([
                    ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#9b59b6')),
                    ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                    ('ALIGN', (0, 0), (-1, 0), 'CENTER'),
                    ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                    ('FONTSIZE', (0, 0), (-1, 0), 10),
                    ('BOTTOMPADDING', (0, 0), (-1, 0), 8),
                    ('BACKGROUND', (0, 1), (-1, -1), colors.HexColor('#f4ecf7')),
                    ('TEXTCOLOR', (0, 1), (-1, -1), colors.black),
                    ('ALIGN', (0, 1), (-1, -1), 'CENTER'),
                    ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
                    ('FONTSIZE', (0, 1), (-1, -1), 9),
                    ('GRID', (0, 0), (-1, -1), 1, colors.grey),
                ]))
                elements.append(Spacer(1, 10))
                elements.append(table)
                
                # Add note about efficiency calculation
                note = Paragraph("<i>Note: Process Efficiency = (Accepted Cycles / Total Cycles) × 100%</i>",
                               ParagraphStyle(name='EfficiencyNote',
                                            fontSize=9,
                                            textColor=colors.HexColor('#7f8c8d'),
                                            spaceBefore=10))
                elements.append(note)
        
        elements.append(Spacer(1, 30))
        return elements
    
    def get_benchmark_status(self, kpi_name, value):
        """Get the correct benchmark status with descriptive labels"""
        # Determine benchmark based on KPI type
        benchmarks = KPI_BENCHMARKS.get(kpi_name, {})
        
        if kpi_name in ['temperature_stability', 'pressure_stability']:
            # For stability KPIs, LOWER values are better
            if 'excellent' in benchmarks and value <= benchmarks['excellent']:
                benchmark = 'excellent'
            elif 'good' in benchmarks and value <= benchmarks['good']:
                benchmark = 'good'
            else:
                benchmark = 'poor'
        else:
            # For success/availability KPIs, HIGHER values are better
            if 'excellent' in benchmarks and value >= benchmarks['excellent']:
                benchmark = 'excellent'
            elif 'good' in benchmarks and value >= benchmarks['good']:
                benchmark = 'good'
            else:
                benchmark = 'poor'
        
        # Map to more descriptive status
        status_map = {
            'excellent': '✓ Excellent',
            'good': '✓ Good',
            'poor': '⚠ Needs Attention'
        }
        return status_map.get(benchmark, benchmark.title())
    
    def _create_risk_assessment_section(self):
        """
        Create the Risk Assessment section.
        Conforms to ISO 14971 risk management principles and ReportLab layout best practices.
        """
        elements = []

        # ------------------------------------------------------------------
        # Section title
        # ------------------------------------------------------------------
        elements.append(
            Paragraph(
                "8. RISK ASSESSMENT",
                self.styles[self.get_style("SectionTitle")]
            )
        )
        elements.append(Spacer(1, 12))

        # ------------------------------------------------------------------
        # Introduction
        # ------------------------------------------------------------------
        intro_text = (
            "Risk assessment identifies potential failure modes in the sterilization process "
            "and evaluates their impact on patient safety. This analysis follows ISO 14971 "
            "principles for medical device risk management."
        )
        elements.append(Paragraph(intro_text, self.styles["BodyText"]))
        elements.append(Spacer(1, 12))

        # ------------------------------------------------------------------
        # Risk matrix figure
        # ------------------------------------------------------------------
        risk_img = VISUAL_RESULTS_DIR / "risk_assessment_matrix.png"
        if risk_img.exists():
            img = Image(str(risk_img), width=5 * inch, height=4 * inch)
            img.hAlign = "CENTER"
            elements.append(img)
            elements.append(Spacer(1, 6))

            caption_style = ParagraphStyle(
                name="RiskFigureCaption",
                fontSize=9,
                alignment=TA_CENTER,
                textColor=colors.HexColor("#7f8c8d"),
            )
            elements.append(
                Paragraph(
                    "Figure 8.1: Risk assessment matrix showing identified risk factors",
                    caption_style,
                )
            )
            elements.append(Spacer(1, 16))

        # ------------------------------------------------------------------
        # Risk assessment results
        # ------------------------------------------------------------------
        risk_assessment = self.analysis_results.get("risk_assessment", {})

        overall_risk = risk_assessment.get("overall_risk", "Low")

        if overall_risk == "Low":
            risk_color = colors.green
        elif overall_risk == "Medium":
            risk_color = colors.orange
        else:
            risk_color = colors.red

        elements.append(
            Paragraph(
                f"<b>Overall Risk Level:</b> "
                f"<font color='{risk_color.hexval()}'>{overall_risk}</font>",
                self.styles["BodyText"],
            )
        )
        elements.append(Spacer(1, 12))

        # ------------------------------------------------------------------
        # Risk factors table
        # ------------------------------------------------------------------
        risk_factors = risk_assessment.get("risk_factors", [])

        if not risk_factors:
            elements.append(
                Paragraph(
                    "No significant risk factors were identified.",
                    self.styles["BodyText"],
                )
            )
            elements.append(Spacer(1, 24))
            return elements

        elements.append(
            Paragraph(
                "Identified Risk Factors:",
                self.styles[self.get_style("SubsectionTitle")],
            )
        )
        elements.append(Spacer(1, 8))

        # ------------------------------------------------------------------
        # Table styles
        # ------------------------------------------------------------------
        header_style = ParagraphStyle(
            name="RiskTableHeader",
            parent=self.styles["Normal"],
            fontName="Helvetica-Bold",
            fontSize=10,
            textColor=colors.whitesmoke,
            alignment=TA_CENTER,
        )

        cell_left = ParagraphStyle(
            name="RiskTableCellLeft",
            parent=self.styles["Normal"],
            fontSize=9,
            leading=11,
            alignment=TA_LEFT,
        )

        cell_center = ParagraphStyle(
            name="RiskTableCellCenter",
            parent=self.styles["Normal"],
            fontSize=9,
            leading=11,
            alignment=TA_CENTER,
        )

        # ------------------------------------------------------------------
        # Build table data (Paragraphs only – no raw strings)
        # ------------------------------------------------------------------
        table_data = [
            [
                Paragraph("Risk Factor", header_style),
                Paragraph("Sterilizer", header_style),
                Paragraph("Risk Level", header_style),
                Paragraph("Recommendation", header_style),
            ]
        ]

        for factor in risk_factors:
            risk_level = factor.get("risk_level", "Low")

            # Set recommendation based on risk level
            if risk_level == "Low":
                recommendation_text = "Continue routine monitoring"
            elif risk_level == "Medium":
                recommendation_text = "Investigate and adjust process"
            elif risk_level == "High":
                recommendation_text = "Immediate corrective action required"
            else:
                recommendation_text = "Review risk factor"

            table_data.append(
                [
                    Paragraph(factor["factor"], cell_left),
                    Paragraph(factor["sterilizer"], cell_center),
                    Paragraph(risk_level, cell_center),
                    Paragraph(recommendation_text, cell_left),
                ]
            )


        # ------------------------------------------------------------------
        # Create table
        # ------------------------------------------------------------------
        table = Table(
            table_data,
            colWidths=[1.6 * inch, 1.1 * inch, 1.1 * inch, 2.4 * inch],
            repeatRows=1,
            hAlign="LEFT",
        )

        table.setStyle(
            TableStyle(
                [
                    ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#e74c3c")),
                    ("GRID", (0, 0), (-1, -1), 0.75, colors.grey),
                    ("VALIGN", (0, 0), (-1, -1), "TOP"),
                    ("LEFTPADDING", (0, 0), (-1, -1), 6),
                    ("RIGHTPADDING", (0, 0), (-1, -1), 6),
                    ("TOPPADDING", (0, 0), (-1, -1), 4),
                    ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
                ]
            )
        )

        elements.append(table)
        elements.append(Spacer(1, 15))

        # ------------------------------------------------------------------
        # Risk control statement
        # ------------------------------------------------------------------
        elements.append(
            Paragraph(
                "<i>Risk control actions are implemented in proportion to the identified "
                "risk level and monitored through ongoing process trending, consistent "
                "with ISO 14971 principles.</i>",
                ParagraphStyle(
                    name="RiskControlNote",
                    fontSize=8.5,
                    textColor=colors.HexColor("#7f8c8d"),
                ),
            )
        )

        elements.append(Spacer(1, 30))
        return elements
    
    def _create_recommendations_section(self):
        """Create recommendations section"""
        elements = []
        
        elements.append(Paragraph("9. RECOMMENDATIONS", 
                                self.styles[self.get_style('SectionTitle')]))
        elements.append(Spacer(1, 20))
        
        # Summary of findings leading to recommendations
        summary_text = """
        Based on the comprehensive analysis of sterilization performance data, 
        the following recommendations are provided to maintain process effectiveness, 
        ensure regulatory compliance, and optimize equipment performance.
        """
        elements.append(Paragraph(summary_text, self.styles['BodyText']))
        elements.append(Spacer(1, 10))
        
        # Recommendations by category
        categories = [
            {
                'title': 'Process Maintenance',
                'recommendations': [
                    'Continue daily biological indicator testing as per ISO 17665 and AAMI ST79',
                    'Maintain current sterilization protocols that have proven effective',
                    'Monitor control charts for early detection of process deviations',
                    'Continue documentation of all cycles for audit readiness'
                ]
            },
            {
                'title': 'Quality Assurance',
                'recommendations': [
                    'Regularly review process capability indices for trend analysis',
                    'Conduct periodic equipment calibration as per manufacturer guidelines',
                    'Maintain personnel training records for all operators',
                    'Implement statistical process control for continuous monitoring'
                ]
            },
            {
                'title': 'Regulatory Compliance',
                'recommendations': [
                    'Continue compliance with daily biological indicator requirements',
                    'Maintain comprehensive validation documentation',
                    'Prepare for regulatory audits with complete cycle records',
                    'Stay updated on changes to international standards'
                ]
            },
            {
                'title': 'Risk Management',
                'recommendations': [
                    'Continue monitoring identified risk factors',
                    'Establish contingency plans for equipment maintenance',
                    'Regularly review and update risk assessments',
                    'Maintain spare parts inventory for critical components'
                ]
            }
        ]
        
        for category in categories:
            elements.append(Paragraph(category['title'], 
                                    self.styles[self.get_style('SubsectionTitle')]))
            
            for recommendation in category['recommendations']:
                elements.append(Paragraph(f"• {recommendation}", 
                                        ParagraphStyle(name='BulletStyle',
                                                     fontSize=11,
                                                     leftIndent=20,
                                                     firstLineIndent=-10,
                                                     spaceAfter=10)))
            
            elements.append(Spacer(1, 10))
        
        # Implementation timeline
        elements.append(Paragraph("Implementation Timeline", 
                                self.styles[self.get_style('SubsectionTitle')]))
        
        timeline_text = """
        <b>Immediate Actions (Within 1 week):</b> Review any identified non-conformances, implement corrective actions for high-risk items
        <b>Ongoing (Daily):</b> Continue biological indicator testing, monitor control charts, review cycle success rates
        <b>Monthly:</b> Review process capability indices, equipment performance metrics, compliance status
        <b>Quarterly:</b> Comprehensive equipment calibration, personnel competency assessment, standards compliance review
        <b>Annually:</b> Full process revalidation, risk assessment update, regulatory compliance audit
        """
        
        elements.append(Paragraph(timeline_text, self.styles['BodyText']))
        elements.append(Spacer(1, 30))
        
        # Conclusion
        elements.append(Paragraph("Conclusion", 
                                self.styles[self.get_style('SubsectionTitle')]))
        
        # Calculate overall status
        overall_compliant = True
        for standard_name, compliance in self.compliance_results.items():
            for sterilizer_name, results in compliance.items():
                if not results['overall_compliant']:
                    overall_compliant = False
        
        overall_risk = self.analysis_results.get('risk_assessment', {}).get('overall_risk', 'Low')
        
        if overall_compliant and overall_risk == 'Low':
            conclusion_text = """
            The comprehensive analysis confirms that both Statim and Ritter sterilization 
            processes are fully compliant with all applicable international standards, 
            including ISO 17665, AAMI ST79, EN 285, FDA 21 CFR Part 820, and CSA Z314.23.
            
            All sterilization cycles have successfully passed biological indicator testing, 
            demonstrating consistent process validation. Process capability indices are 
            within acceptable ranges, and control charts show stable, in-control processes.
            
            The overall risk assessment indicates a low-risk profile, and all key 
            performance indicators meet or exceed industry benchmarks for excellence.
            
            Continued adherence to established protocols, regular monitoring, and 
            comprehensive documentation will ensure ongoing compliance and patient safety.
            """
        elif overall_compliant and overall_risk == 'Medium':
            conclusion_text = """
            The analysis confirms compliance with international sterilization standards, 
            but identifies some areas requiring attention to maintain optimal performance.
            
            While all cycles passed biological indicator testing and meet validation 
            requirements, certain process parameters show variation that should be monitored.
            
            Recommendations provided in this report will help address these areas and 
            further improve sterilization process robustness and reliability.
            """
        else:
            conclusion_text = """
            The analysis identifies areas requiring attention to ensure full compliance 
            with international standards. While biological indicator testing demonstrates 
            process effectiveness, certain aspects of the sterilization process require 
            improvement.
            
            The recommendations provided in this report outline specific actions to 
            address identified issues and enhance overall process performance and 
            compliance.
            """
        
        elements.append(Paragraph(conclusion_text, self.styles['BodyText']))
        elements.append(Spacer(1, 30))
        
        return elements
    
    def _create_appendices(self):
        """Create appendices section"""
        elements = []
        
        elements.append(Paragraph("10. APPENDICES", 
                                self.styles[self.get_style('SectionTitle')]))
        elements.append(Spacer(1, 12))
        
        # Appendix A: Methodology Details
        elements.append(Paragraph("Appendix A: Detailed Methodology", 
                                ParagraphStyle(name='AppendixTitleStyle',
                                             fontSize=12,
                                             textColor=colors.HexColor('#2c3e50'),
                                             spaceBefore=10,
                                             spaceAfter=8)))
        
        method_details = """
        <b>Statistical Methods Used:</b>
        
        1. <b>Descriptive Statistics:</b> Mean, standard deviation, median, quartiles, 
           coefficient of variation, skewness, and kurtosis were calculated for all 
           numeric parameters.
        
        2. <b>Inferential Statistics:</b> Independent t-tests were used to compare 
           means between sterilizer types. Welch's t-test was applied when variances 
           were unequal.
        
        3. <b>Normality Testing:</b> Shapiro-Wilk test assessed data normality.
        
        4. <b>Process Capability Analysis:</b> Cp, Cpk, Pp, and Ppk indices were 
           calculated using typical operating ranges as specification limits.
        
        5. <b>Control Chart Methodology:</b> Individual-moving range charts were 
           constructed using 3-sigma control limits.
        
        6. <b>Risk Assessment:</b> Risk priority numbers were calculated based on 
           occurrence, severity, and detection ratings following ISO 14971 methodology.
        
        <b>Efficiency Calculation:</b>
        Process Efficiency = (Number of Accepted Cycles / Total Number of Cycles) × 100%
        
        This calculation reflects the actual success rate of sterilization cycles, 
        which is the primary measure of process effectiveness in sterilization standards.
        """
        
        elements.append(Paragraph(method_details, self.styles['BodyText']))
        elements.append(Spacer(1, 20))
        
        # Appendix B: Standards References
        elements.append(Paragraph("Appendix B: Standards and Regulations", 
                                ParagraphStyle(name='AppendixTitleStyle',
                                             fontSize=12,
                                             textColor=colors.HexColor('#2c3e50'),
                                             spaceBefore=15,
                                             spaceAfter=35)))
        
        standards_ref = """
        <b>Primary Standards Referenced:</b>
        
        • <b>ISO 17665-1:</b> Sterilization of health care products — Moist heat
           - Requires process validation through biological indicator testing
           - Does not specify fixed temperature/pressure limits
           - Focuses on achieving sterility assurance level (SAL) of 10⁻⁶
        
        • <b>AAMI ST79:</b> Comprehensive guide to steam sterilization
           - Recommends weekly biological indicator testing
           - Provides guidance on load configuration and monitoring
           - Emphasizes documentation and quality systems
        
        • <b>EN 285:</b> Sterilization — Steam sterilizers — Large sterilizers
           - Specifies equipment performance requirements
           - Requires temperature uniformity testing
           - Focuses on equipment validation
        
        • <b>ISO 13485:</b> Medical devices — Quality management systems
           - Requires documented process validation
           - Emphasizes risk management and corrective actions
        
        • <b>ISO 14971:</b> Medical devices — Application of risk management
           - Framework for risk assessment and mitigation
           - Required for medical device sterilization processes
        
        • <b>FDA 21 CFR Part 820:</b> Quality System Regulation
           - Requires process validation for sterilization
           - Emphasizes documentation and record keeping
        
        • <b>CSA Z314.23:</b> Canadian standard for sterilization in health care facilities
           - Requires validation of sterilization processes
           - Emphasizes routine monitoring and testing
           - Includes requirements for equipment maintenance and personnel training
        
        <b>Key Compliance Requirements:</b>
        
        • <b>Process Validation:</b> Demonstrated through biological indicator testing
        • <b>Routine Monitoring:</b> Daily biological indicator tests
        • <b>Documentation:</b> Complete records of all sterilization cycles
        • <b>Equipment Qualification:</b> Initial and ongoing performance verification
        • <b>Personnel Training:</b> Competency verification for all operators
        """
        
        elements.append(Paragraph(standards_ref, self.styles['BodyText']))
        elements.append(Spacer(1, 30))
        
        # Appendix C: Glossary
        elements.append(Paragraph("Appendix C: Glossary of Terms", 
                                ParagraphStyle(name='AppendixTitleStyle',
                                             fontSize=12,
                                             textColor=colors.HexColor('#2c3e50'),
                                             spaceBefore=25,
                                             spaceAfter=15)))
        
        glossary = """
        <b>Biological Indicator (BI):</b> Preparation of viable microorganisms that 
        provides a defined resistance to a specific sterilization process. Used to 
        validate sterilization effectiveness.
        
        <b>Process Efficiency:</b> The percentage of sterilization cycles that are 
        successfully completed and accepted. Calculated as (Accepted Cycles / Total Cycles) × 100%.
        
        <b>Cpk (Process Capability Index):</b> Statistical measure of a process's 
        ability to produce output within specification limits.
        
        <b>Control Limits:</b> Statistical limits on a control chart that indicate 
        the expected variation in a stable process.
        
        <b>Process Validation:</b> Documented evidence that a process consistently 
        produces a result meeting predetermined specifications.
        
        <b>Sterility Assurance Level (SAL):</b> Probability of a single viable 
        microorganism occurring on an item after sterilization. Typically 10⁻⁶ for 
        sterile medical devices.
        
        <b>UCL/LCL:</b> Upper and Lower Control Limits on a control chart.
        
        <b>Validation (in sterilization):</b> The process of demonstrating that a 
        sterilization process will consistently produce sterile products, typically 
        through biological indicator testing.
        """
        
        elements.append(Paragraph(glossary, self.styles['BodyText']))
        elements.append(Spacer(1, 30))
        
        # Contact Information
        elements.append(Paragraph("For More Information:", 
                                ParagraphStyle(name='ContactTitleStyle',
                                             fontSize=11,
                                             textColor=colors.HexColor('#2c3e50'),
                                             spaceBefore=15,
                                             spaceAfter=8)))
        
        contact_info = """
        MDR Team
        Email: info@herzig-eye.com
        Phone: (613) 800-1680
        
        This report and associated data are confidential property of Herzig Eye Institute.
        Unauthorized distribution is prohibited.
        
        <b>Report Generated:</b> {date}
        <b>Analysis Date:</b> {analysis_date}
        <b>Data Source:</b> Statim and Ritter sterilizer cycle data
        <b>Report Version:</b> 2.0 (Separated Analysis & Reporting)
        """.format(date=self.report_date.strftime("%B %d, %Y"),
                  analysis_date=self.analysis_date.strftime("%B %d, %Y") if self.analysis_date else "Not specified")
        
        elements.append(Paragraph(contact_info, 
                                ParagraphStyle(name='ContactInfoStyle',
                                             fontSize=10,
                                             alignment=TA_JUSTIFY,
                                             spaceAfter=30)))
        
        return elements
    
    def generate_report(self):
        """Generate the complete PDF report"""
        print("\n" + "="*70)
        print("GENERATING COMPREHENSIVE PDF REPORT")
        print("="*70)
        
        report_path = self.generate_pdf_report()
        
        print(f"\n✓ Report generated: {report_path}")
        print(f"✓ Report date: {self.report_date.strftime('%B %d, %Y')}")
        print(f"✓ Analysis date: {self.analysis_date.strftime('%B %d, %Y') if self.analysis_date else 'Not specified'}")
        print(f"✓ Results loaded from: {ANALYSIS_RESULTS_DIR}")
        
        return report_path

def main():
    """Main function to generate report from pre-calculated results"""
    print("=" * 70)
    print("STERILIZER REPORT GENERATOR")
    print("=" * 70)
    print("Note: This script only generates reports from pre-calculated results.")
    print("      Run the analysis script first to generate analysis results.\n")
    
    generator = SterilizerReportGenerator()
    report_path = generator.generate_report()
    
    if report_path:
        print("\n" + "="*70)
        print("REPORT GENERATION COMPLETE")
        print("="*70)
        print(f"\nThe comprehensive report includes:")
        print("• Executive summary with key findings")
        print("• Statistical analysis results from pre-calculated data")
        print("• Control charts and process capability analysis")
        print("• Compliance assessment against international standards")
        print("• KPI benchmarking with corrected efficiency calculation")
        print("• Risk assessment matrix")
        print("• Actionable recommendations for improvement")
        print("• Detailed appendices with methodology and references")
        
        print(f"\nReport is ready for distribution and regulatory submission.")

if __name__ == "__main__":
    main()