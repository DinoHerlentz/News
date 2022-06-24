import requests

def BBC():
    query_params = {
        "source": "bbc-news",
        "sortBy": "top",
        "apiKey": "05efc3599af34ed790f971025ff3f3ef"
    }
    main_url = "https://newsapi.org/v1/articles"

    res = requests.get(main_url, params=query_params)
    bbc_page = res.json()

    article = bbc_page['articles']
    results = []

    for ar in article:
        results.append(ar['title'])
    
    for i in range(len(results)):
        print(i + 1, results[i])
    
    from win32com.client import Dispatch
    speak = Dispatch("SAPI.Spvoice")
    speak.Speak(results)

if __name__ == "__main__":
    BBC()