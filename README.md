## How to use:
```
# clone this repo
git clone https://github.com/Wangsherpa/Excel2Json.git

# install required libraries
pip install -r requirements.txt

# json file will not be saved
python convert.py --source_path "data.xlsx" --sheet_name "Marks" --orient "index"

# json file will be saved as "new_output.json"
python convert.py --source_path "data.xlsx" --sheet_name "Marks" --orient "index" --target_path "new_output.json"
```
