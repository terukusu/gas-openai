# gas-openai
- GAS(Google Apps Script)から使えるOpenAIクライアントライブラリです。

## 特徴
- 超シンプルなインターフェイス
- 流量制限時の自動リトライ対応
- テキストを生成・・OK！
- JSON出力・・OK！
- 画像分析・・OK！
- 画像生成・・OK！
- 音声文字起こし・・OK！
- 関数呼び出しによる前提知識補完・・OK！
- エンベディング・・・OK！
- Open AI とAzure 両方に対応
- デフォルトでも動くけど、パラメータで柔軟にカスタマイズも可能(モデル名とか)

# 使い方
- [src/Code.js](src/Code.js) の内容を何処かのGASへ保存
- 使いたいGASへライブラリとして追加する
    - スクリプトエディタの左メニューの「ライブラリ ＋」の ＋ をクリックして、↑のスクリプトIDを指定
- 詳しい使い方は、[src/Code.js](src/Code.js) や↓のサンプルを斜め読みしてください

# シンプルなコードの例
```JavaScript
//OpenAIの場合
const client = OpenAI.createOpenAIClient({
  apiKey: '<YOUR_API_KEY>'
});

// Azure の場合
const client = OpenAI.createAzureOpenAIClient({
  apiKey: '<YOUR_API_KEY>',
  azureEndpoint: '<AZURE_ENDPOINT>',
});

const response = client.simpleChat("こんにちは!");
```

# AIの回答をJSONで受け取る例
```JavaScript
  // ==== AIの回答をJSONで受け取る例 ====
  const responseSchemaHuman = {
    "title": "Human",
    "type": "object",
    "properties": {
      "name": {
        "type": "string",
        "description": "名前"
      },
      "age": {
        "type": "number",
        "description": "年齢"
      }
    },
    "required": ["name", "age"]
  };
  
  params = {
    responseSchema: responseSchemaHuman
  }

  result = client.simpleChat("架空の人物になって自己紹介をして", params);
  Logger.log(result);
  // 出力例：
  // {
  //   "name": "ヴィクトリア",
  //   "age": 27
  // }
  //
```



JSONスキーマは、受け取りたいJSONっぽい雰囲気のものを書いて、以下のようにChatGPTにJSONスキーマにしてもらえばOK。  
<img width="640" src="https://github.com/terukusu/gas-openai/assets/205033/a4dafd55-007a-406f-b4b9-2cf960c295a8">

# AIの前提知識を関数呼び出しで補完する例
```JavaScript
  // ==== AIに利用可能な関数を伝えて、必要に応じて実行させる例(Function Calling機能) ====
  function getWeather(location) {
    return {weather: "晴れ"};
  }
  
  const functions = [{
    func: getWeather,
    description: "指定された地域の天気を調べます。",
    parameters: {
      "$schema": "http://json-schema.org/draft-07/schema#",
      "type": "object",
      "properties": {
        "loation": {
          "type": "string",
          "description": "天気を調べたい地域"
        }
      },
      "required": ["location"],
      "additionalProperties": false
    }
  }];
  
  params = {
    functions: functions,
  }
  
  result = client.simpleChat("アラスカの天気は？", params);
  Logger.log(result);
  // 出力例：
  // アラスカの天気は晴れです。
  //
````

# 画像をAIで処理する例
```JavaScript
  // ==== Drive上の画像を処理する例 ====
  const myPng  = DriveApp.getFileById("1KRr_7CdjYklHwSvL7EfRfp0EiStgxTIq");

  params = {
    model: "gpt-4-vision-preview", // 画像を使うときはこのモデルを指定
    images:[myPng.getBlob()]
  };

  result = client.simpleChat("この画像を解説してください。", params);
  Logger.log(result);
  // 出力例：
  // この画像は、スマートフォンやタブレットなどのデバイスで使用されるメニューの一部を示しています。画面の左側には、各メニュー項目のアイコンが表示されており、右側にはその項目の名前が書かれています。
  //
```

# AIで画像を生成する例
```JavaScript
  params = {
    model: "dall-e-3", // 画像生成を使うときはこのモデルを指定
  };

  result = client.simpleImageGeneration("犬を描いてください。", params);
  Logger.log(result);
  // 出力例：
  // https://the.url.of/generated/image
  // (生成された画像のURLが出力される）
  //
```

# 音声の文字起こしをする例
```JavaScript
  // ==== 音声の文字起こしの例 ====
  const myVoice  = DriveApp.getFileById("1nPivg4JwHhrE4Qax1du2CM2P5uIIY9Py");

  params = {
    model: "whisper" // Azure の whisper
    // model: "whisper-1" // 本家 OpenAI の whisper
  };

  result = client.simpleVoiceToText(myVoice.getBlob(), params);
  Logger.log(result);
  // 出力例：
  // みなさんお集まりいただきありがとうございます 本日は当プロジェクトの大型アップデートについて話し合いましょう
```

# エンベディング（文字列のベクトル表現化）をする例
```JavaScript
  // ==== エンベディング（ベクトル化）の例 ====
  params = {
    model: "text-embedding-3-small", // embeddingsを使うときはこのモデルを指定。text-embedding-3-large でも良いよ。
  };

  result = client.simpleEmbedding(["わーい"], params);
  Logger.log(result);
```
