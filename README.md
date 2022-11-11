# my_webdrivers

A Python library for custom data connections.

## Installation

pip install git+https://github.com/edire/my_excel.git

## Usage

```python
import my_excel

with my_excel.Excel('.\file_path.xlsx', import_vba=True, visible=True) as xl:
	xl.refresh_all()
	xl.save()
	xl.close(save=True)
```

## Contributing

Pull requests are welcome. For major changes, please open an issue first to discuss what you would like to change.

## License

MIT License

## Release Updates

Adjusted Excel workbook counter before closing application.