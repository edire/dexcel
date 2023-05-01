# Description

A Python library by Dire Analytics for controlling Microsoft Excel files.

## Installation

pip install git+https://github.com/edire/dexcel.git

## Usage

```python
import dexcel

with dexcel.Excel('.\file_path.xlsx', import_vba=True, visible=True) as xl:
	xl.refresh_all()
	xl.save()
	xl.close(save=True)
```

## Contributing

Pull requests are welcome. For major changes, please open an issue first to discuss what you would like to change.

## License

MIT License

## Updates

05/01/2023 - Added apostrophes around workbook name in run macro command.