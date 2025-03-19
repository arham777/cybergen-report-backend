from main import app

# This is the entry point that Render will look for
if __name__ == "__main__":
    import uvicorn
    uvicorn.run("main:app", host="0.0.0.0", port=8000) 