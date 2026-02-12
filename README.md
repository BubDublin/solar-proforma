# ⚡ Solar Pro-Forma Generator

A professional web application for generating solar project pro-formas with live calculations.

## Features

- ✅ **Live Preview**: See key metrics update as you type
- ✅ **4 Toggles**: ITC, SREC Program, Utility, Escalation Rate
- ✅ **17 Cost Categories**: Full pricing breakdown
- ✅ **25-Year Cash Flow**: Year-by-year projections
- ✅ **Client Summary**: Professional presentation view
- ✅ **Excel Export**: Download complete pro-forma with formulas

## Market Data Included

- **DC SREC**: $380-455 (2025), ACP schedule through 2042
- **MD SREC**: $48-50 standard, $70-74 Brighter Tomorrow
- **Utility Rates**: PEPCO MD/DC, BGE, Potomac Edison

## Deployment Instructions

### Option 1: Streamlit Cloud (FREE - Recommended)

1. Create a free account at [streamlit.io](https://streamlit.io)
2. Create a new app and connect your GitHub repository
3. Or deploy directly:
   - Go to [share.streamlit.io](https://share.streamlit.io)
   - Click "New app"
   - Select your repo containing these files
   - Main file path: `app.py`
4. Your app will be live at: `https://your-app-name.streamlit.app`

### Option 2: Local Development

```bash
# Install dependencies
pip install -r requirements.txt

# Run the app
streamlit run app.py
```

The app will open in your browser at `http://localhost:8501`

## Usage

1. **Enter Project Info**: Customer name, project name, system size
2. **Select Location**: Maryland or DC
3. **Set Toggles**: ITC status, SREC program, utility, escalation
4. **Enter Pricing**: $/W for each cost category
5. **View Live Preview**: See payback, Year 1 benefits, etc.
6. **Generate Excel**: Click button to download complete pro-forma

## Excel Output

The generated Excel file includes:
- **Inputs & Assumptions**: All your inputs and toggles
- **25-Year Cash Flow**: Year-by-year projections with formulas
- **Client Summary**: Professional presentation view

## Customization

Edit `app.py` to:
- Change default pricing values
- Add more cost categories
- Modify utility rates
- Adjust SREC ACP schedules
- Change styling/colors

## Support

Built for Captain Power Solar 🚀
