from fastapi import FastAPI, Request
import pandas as pd


app = FastAPI()


@app.get("/excel")
def get_excel():
    """Get all the data from the excel"""
    df = pd.read_excel('data.xlsx', sheet_name='data')
    return df.to_json()


@app.post("/excel")
def update_excel(request):
    df = pd.read_excel('data.xlsx', sheet_name='data')
    return {'message': 'ok'}


@app.delete("/excel")
def delete_excel():
    """Delete all the data from excel"""
    df = pd.DataFrame()
    df.to_excel('data.xlsx', sheet_name='data')
    return {'message': 'ok'}


@app.put("/excel")
async def put_excel(request: Request):
    """Create new column in excel"""
    df = pd.read_excel('data.xlsx', sheet_name='data')
    data = await request.json()
    new_col = data.get('col')
    df[new_col] = pd.Series()
    df.to_excel('data.xlsx', sheet_name='data')
    return {'message': 'ok'}


@app.patch("/excel")
async def patch_excel(request: Request):
    """Rename a column in excel"""
    try:
        df = pd.read_excel('data.xlsx', sheet_name='data')
        data = await request.json()
        col = data.get('col')
        new_col = data.get('new col')
        df.rename(columns={col: new_col}, inplace=True)
        df.to_excel('data.xlsx', sheet_name='data')
        return {'message': 'ok'}
    except Exception as e:
        return {'message': 'error'}

