# my_webdrivers

A Python library for custom data connections.

## Installation

pip install git+https://github.com/edire/my_excel.git

## Usage

```python
import my_excel

xl = Excel('.\file_path.xlsx', visible=True)
xl.refresh_all()
xl.save()
xl.close()
```

## Contributing

Pull requests are welcome. For major changes, please open an issue first to discuss what you would like to change.

## License

MIT License