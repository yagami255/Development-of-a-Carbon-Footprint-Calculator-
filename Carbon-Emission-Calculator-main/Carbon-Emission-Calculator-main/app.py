from flask import Flask, request, render_template, send_file, session
import io
import csv
import json
from datetime import datetime
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
from reportlab.lib.colors import HexColor
from reportlab.lib.units import inch

app = Flask(__name__)
app.secret_key = 'your_secret_key_change_in_production'  # Change this!

# Emission Factors
EMISSION_FACTORS = {
    'electricity_bangladesh': 0.62,
    'diesel': 3.12,
    'petrol': 2.66,
    'cng': 1.95,
    'electronics_spend': 0.25,
    'cloud_spend': 0.20,
    'flight_economy': 0.15,
    'flight_business': 0.45,
    'commute_bus': 0.08,
    'commute_car': 0.17,
    'commute_cng_rickshaw': 0.10,
    'commute_rickshaw': 0.0,
    'refrigerant_gwp': {
        'R-410A': 2088,
        'R-22': 1810,
    }
}

# Calculation Functions
def calculate_scope1(generator_fuel_liters, generator_fuel_type, refrigerant_kg, refrigerant_type, vehicle_fuel_volume, vehicle_fuel_type):
    gen_emissions = generator_fuel_liters * EMISSION_FACTORS.get(generator_fuel_type.lower(), 0)
    ref_emissions = refrigerant_kg * EMISSION_FACTORS['refrigerant_gwp'].get(refrigerant_type.upper(), 0)
    if vehicle_fuel_type.lower() == 'cng':
        veh_emissions = vehicle_fuel_volume * EMISSION_FACTORS['cng']
    else:
        veh_emissions = vehicle_fuel_volume * EMISSION_FACTORS.get(vehicle_fuel_type.lower(), 0)
    return gen_emissions + ref_emissions + veh_emissions

def calculate_scope2(electricity_kwh):
    return electricity_kwh * EMISSION_FACTORS['electricity_bangladesh']

def calculate_scope3(electronics_spend_usd, cloud_spend_usd, flights, headcount, commute_modes, avg_commute_distance_km, total_wfh_days, workdays=250):
    goods_emissions = electronics_spend_usd * EMISSION_FACTORS['electronics_spend']
    cloud_emissions = cloud_spend_usd * EMISSION_FACTORS['cloud_spend']
    
    flights_emissions = 0
    for flight in flights:
        dist_km, class_type = flight
        factor = EMISSION_FACTORS['flight_' + class_type.lower()]
        flights_emissions += dist_km * factor
    
    commute_emissions = 0
    effective_workdays = workdays * (1 - (total_wfh_days / (headcount * workdays))) if headcount > 0 else 0
    for mode, percentage in commute_modes.items():
        mode_factor = EMISSION_FACTORS.get('commute_' + mode.lower(), 0)
        pkm = headcount * (percentage / 100) * avg_commute_distance_km * 2 * effective_workdays
        commute_emissions += pkm * mode_factor
    
    return goods_emissions + cloud_emissions + flights_emissions + commute_emissions

def generate_csv(report_data):
    output = io.StringIO()
    writer = csv.writer(output)
    writer.writerow(['Carbon Footprint Report', datetime.now().strftime('%Y-%m-%d %H:%M:%S')])
    writer.writerow([])
    writer.writerow(['Category', 'Subcategory', 'Value', 'Unit', 'Emissions (kg CO2e)'])
    
    for section, items in report_data.items():
        writer.writerow([section, '', '', '', ''])
        for item in items:
            writer.writerow(['', item['name'], item['value'], item['unit'], f"{item['emissions']:.2f}"])
        writer.writerow(['', 'Total ' + section, '', '', f"{sum(i['emissions'] for i in items):.2f}"])
        writer.writerow([])
    
    total_emissions = sum(sum(i['emissions'] for i in items) for items in report_data.values())
    writer.writerow(['Total Emissions', '', '', '', f"{total_emissions:.2f} kg CO2e"])
    
    output.seek(0)
    return output

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        session['company_name'] = request.form.get('company_name', 'Company')
        return render_template('calculator.html')
    return render_template('index.html')

@app.route('/calculate', methods=['POST'])
def calculate():
    headcount = float(request.form.get('headcount', 0))
    electricity_kwh = float(request.form.get('electricity_kwh', 0))
    generator_fuel_liters = float(request.form.get('generator_fuel_liters', 0))
    generator_fuel_type = request.form.get('generator_fuel_type', 'diesel')
    refrigerant_kg = float(request.form.get('refrigerant_kg', 0))
    refrigerant_type = request.form.get('refrigerant_type', 'R-410A')
    owns_vehicles = request.form.get('owns_vehicles', 'no') == 'yes'
    vehicle_fuel_volume = float(request.form.get('vehicle_fuel_volume', 0)) if owns_vehicles else 0
    vehicle_fuel_type = request.form.get('vehicle_fuel_type', 'diesel') if owns_vehicles else ''
    electronics_spend_usd = float(request.form.get('electronics_spend_usd', 0))
    cloud_spend_usd = float(request.form.get('cloud_spend_usd', 0))
    num_flights = int(request.form.get('num_flights', 0))
    flights = []
    for i in range(num_flights):
        dist = float(request.form.get(f'flight_dist_{i}', 0))
        cls = request.form.get(f'flight_class_{i}', 'economy')
        if dist > 0:
            flights.append((dist, cls))
    commute_modes = {
        'bus': float(request.form.get('commute_bus', 0)),
        'cng_rickshaw': float(request.form.get('commute_cng_rickshaw', 0)),
        'rickshaw': float(request.form.get('commute_rickshaw', 0)),
        'car': float(request.form.get('commute_car', 0)),
    }
    avg_commute_distance_km = float(request.form.get('avg_commute_distance_km', 0))
    total_wfh_days = float(request.form.get('total_wfh_days', 0))

    scope1 = calculate_scope1(generator_fuel_liters, generator_fuel_type, refrigerant_kg, refrigerant_type, vehicle_fuel_volume, vehicle_fuel_type)
    scope2 = calculate_scope2(electricity_kwh)
    scope3 = calculate_scope3(electronics_spend_usd, cloud_spend_usd, flights, headcount, commute_modes, avg_commute_distance_km, total_wfh_days)
    total = scope1 + scope2 + scope3
    per_employee = total / headcount if headcount > 0 else 0

    report_data = {
        'Scope 1': [
            {'name': 'Generator Fuel', 'value': generator_fuel_liters, 'unit': 'liters/m³', 'emissions': generator_fuel_liters * EMISSION_FACTORS.get(generator_fuel_type.lower(), 0)},
            {'name': 'Refrigerants', 'value': refrigerant_kg, 'unit': 'kg', 'emissions': refrigerant_kg * EMISSION_FACTORS['refrigerant_gwp'].get(refrigerant_type.upper(), 0)},
            {'name': 'Company Vehicles', 'value': vehicle_fuel_volume, 'unit': 'liters/m³', 'emissions': vehicle_fuel_volume * (EMISSION_FACTORS['cng'] if vehicle_fuel_type.lower() == 'cng' else EMISSION_FACTORS.get(vehicle_fuel_type.lower(), 0))}
        ],
        'Scope 2': [
            {'name': 'Purchased Electricity', 'value': electricity_kwh, 'unit': 'kWh', 'emissions': scope2}
        ],
        'Scope 3': [
            {'name': 'Hardware Purchases', 'value': electronics_spend_usd, 'unit': 'USD', 'emissions': electronics_spend_usd * EMISSION_FACTORS['electronics_spend']},
            {'name': 'Cloud Services', 'value': cloud_spend_usd, 'unit': 'USD', 'emissions': cloud_spend_usd * EMISSION_FACTORS['cloud_spend']},
            {'name': 'Business Flights', 'value': sum(f[0] for f in flights), 'unit': 'passenger-km', 'emissions': sum(f[0] * EMISSION_FACTORS['flight_' + f[1].lower()] for f in flights)},
            {'name': 'Employee Commuting (adjusted for WFH)', 'value': 'N/A', 'unit': '', 'emissions': scope3 - (electronics_spend_usd * EMISSION_FACTORS['electronics_spend'] + cloud_spend_usd * EMISSION_FACTORS['cloud_spend'] + sum(f[0] * EMISSION_FACTORS['flight_' + f[1].lower()] for f in flights))}
        ]
    }

    company_name = session.get('company_name', 'Company')

    return render_template('results.html',
                           company_name=company_name,
                           scope1=scope1,
                           scope2=scope2,
                           scope3=scope3,
                           total=total,
                           per_employee=per_employee,
                           report_data=report_data)

@app.route('/export_csv', methods=['POST'])
def export_csv():
    report_data = json.loads(request.form['report_data'])
    csv_output = generate_csv(report_data)
    return send_file(
        io.BytesIO(csv_output.getvalue().encode('utf-8')),
        mimetype='text/csv',
        as_attachment=True,
        download_name='carbon_footprint_report.csv'
    )

@app.route('/export_pdf', methods=['POST'])
def export_pdf():
    try:
        report_data = json.loads(request.form['report_data'])
        company_name = request.form.get('company_name', 'Company')
        scope1 = float(request.form.get('scope1', 0))
        scope2 = float(request.form.get('scope2', 0))
        scope3 = float(request.form.get('scope3', 0))
        total = float(request.form.get('total', 0))
        per_employee = float(request.form.get('per_employee', 0))

        buffer = io.BytesIO()
        doc = SimpleDocTemplate(buffer, pagesize=A4, rightMargin=40, leftMargin=40, topMargin=50, bottomMargin=50)
        styles = getSampleStyleSheet()
        
        styles.add(ParagraphStyle(name='CompanyName', fontSize=18, alignment=1, spaceAfter=10))
        styles.add(ParagraphStyle(name='ReportTitle', fontSize=24, alignment=1, spaceAfter=20, textColor=colors.darkgreen))
        styles.add(ParagraphStyle(name='Date', fontSize=12, alignment=1, spaceAfter=40))
        styles.add(ParagraphStyle(name='DetailHeader', fontSize=16, spaceBefore=40, spaceAfter=15, textColor=colors.darkgreen))

        story = []

        # Header
        story.append(Paragraph(company_name, styles['CompanyName']))
        story.append(Paragraph("Carbon Footprint Results", styles['ReportTitle']))
        story.append(Paragraph(f"Generated on {datetime.now().strftime('%B %d, %Y')}", styles['Date']))

     

        # Detailed Breakdown
        story.append(Paragraph("Detailed Breakdown", styles['DetailHeader']))

        table_data = [['Category', 'Subcategory', 'Value', 'Unit', 'Emissions (kg CO₂e)']]
        row_index = 1

        for section, items in report_data.items():
            table_data.append([section, '', '', '', ''])
            row_index += 1
            for item in items:
                table_data.append(['', item['name'], str(item['value']), item['unit'], f"{item['emissions']:.2f}"])
                row_index += 1
            section_total = sum(i['emissions'] for i in items)
            table_data.append(['', 'Total ' + section, '', '', f"{section_total:.2f}"])
            row_index += 1

        table_data.append(['Total Emissions', '', '', '', f"{total:.2f}"])

        detail_table = Table(table_data, colWidths=[100, 180, 80, 70, 120])
        detail_table.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,0), colors.grey),
            ('TEXTCOLOR', (0,0), (-1,0), colors.white),
            ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
            ('ALIGN', (0,0), (-1,0), 'CENTER'),
            ('GRID', (0,0), (-1,-1), 0.5, colors.lightgrey),

            # Section headers
            ('SPAN', (0,1), (4,1)),
            ('SPAN', (0,5), (4,5)),
            ('SPAN', (0,9), (4,9)),
            ('BACKGROUND', (0,1), (4,1), colors.lightgrey),
            ('BACKGROUND', (0,5), (4,5), colors.lightgrey),
            ('BACKGROUND', (0,9), (4,9), colors.lightgrey),

            # Section totals
            ('BACKGROUND', (0,4), (-1,4), colors.lightgreen),
            ('BACKGROUND', (0,6), (-1,6), colors.lightgreen),
            ('BACKGROUND', (0,12), (-1,12), colors.lightgreen),

            # Final total
            ('BACKGROUND', (0,-1), (-1,-1), colors.black),
            ('TEXTCOLOR', (0,-1), (-1,-1), colors.white),
            ('FONTNAME', (0,-1), (-1,-1), 'Helvetica-Bold'),
            ('SPAN', (0,-1), (3,-1)),
            ('ALIGN', (4,-1), (4,-1), 'RIGHT'),

            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ]))
        story.append(detail_table)

        doc.build(story)
        buffer.seek(0)

        filename = f"{company_name.replace(' ', '_')}_Carbon_Report_{datetime.now().strftime('%Y%m%d')}.pdf"
        return send_file(buffer, mimetype='application/pdf', as_attachment=True, download_name=filename)

    except Exception as e:
        print(f"PDF Error: {e}")
        import traceback
        traceback.print_exc()
        return "Error generating PDF.", 500

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=10000, debug=True)
