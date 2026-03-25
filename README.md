# test_reposu

This is only for test.

---

## Service Invoice Sample

A sample Word document service invoice (`invoice.docx`) is included in this repository.

### Invoice details

| Day | Description | Hours | Rate | Amount |
|-----|-------------|------:|-----:|-------:|
| 1 | Professional Services | 9.0 | $60.00 | $540.00 |
| 2 | Professional Services | 6.5 | $60.00 | $390.00 |
| | **Total** | **15.5** | | **$930.00** |

### Download

Download **[invoice.docx](invoice.docx)** directly from the repository.

### Regenerate the invoice

If you need to customise the invoice (client name, dates, line items, rates, etc.)
you can edit and re-run the generator script:

```bash
# Install dependency (one-time)
pip install python-docx

# Generate invoice.docx
python generate_invoice.py
```

The script `generate_invoice.py` at the root of the repository contains all
invoice data and styling in one place.
