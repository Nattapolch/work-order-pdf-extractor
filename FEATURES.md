# New Features: Token Counting & Cost Calculation

## Overview
Added comprehensive token counting and cost calculation features with support for the latest GPT-4.1 model family (2025).

## New Features

### 🤖 Multiple Model Support
- **GPT-4.1 Nano** (Default): Fastest and cheapest model - $0.10/$0.40 per 1M tokens
- **GPT-4.1 Mini**: Balanced performance - $0.40/$1.60 per 1M tokens  
- **GPT-4.1**: Most capable model - $3.00/$12.00 per 1M tokens

All models support up to 1 million token context window with June 2024 knowledge cutoff.

### 💰 Real-time Cost Tracking
- **Token Usage**: Input/output token counting for each API call
- **Cost Calculation**: Automatic USD and Thai Baht conversion (33 THB = 1 USD)
- **Session Statistics**: Cumulative tracking across all API calls
- **Live Updates**: Real-time cost display during processing

### 📊 Enhanced UI
- **Model Selection**: Dropdown with pricing information
- **Cost Dashboard**: Live session statistics in Settings tab
- **Progress Display**: Current session cost shown during processing
- **Reset Function**: Clear session statistics with one click

## Cost Information (2025 Rates)

### GPT-4.1 Nano
- **Input**: $0.10 per million tokens (฿3.30)
- **Output**: $0.40 per million tokens (฿13.20)
- **Best for**: High-volume, cost-sensitive applications

### GPT-4.1 Mini  
- **Input**: $0.40 per million tokens (฿13.20)
- **Output**: $1.60 per million tokens (฿52.80)
- **Best for**: Balanced performance and cost

### GPT-4.1
- **Input**: $3.00 per million tokens (฿99.00)
- **Output**: $12.00 per million tokens (฿396.00)
- **Best for**: Complex tasks requiring maximum capability

## Usage Example

Processing 100 work order PDFs with GPT-4.1 Nano:
- Average input: ~2,000 tokens per image
- Average output: ~50 tokens per response
- **Total cost**: ~$0.022 USD (฿0.73 THB) for 100 files

## Benefits

1. **Cost Transparency**: Know exactly how much each processing session costs
2. **Budget Planning**: Track usage patterns for budget forecasting
3. **Model Optimization**: Compare costs across different models
4. **Thai Baht Support**: Local currency display for easier budgeting
5. **Real-time Monitoring**: Live cost updates during processing

## Technical Implementation

- **Token Tracking**: Uses OpenAI API response usage data
- **Cost Calculation**: Real-time calculation with current pricing
- **Currency Conversion**: Automatic USD to THB conversion
- **Session Persistence**: Statistics maintained throughout application session
- **Model Validation**: Ensures selected model is supported

## Future Enhancements

- Export cost reports to CSV
- Historical usage tracking
- Budget alerts and limits
- Batch processing cost estimates