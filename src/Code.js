// LLM関連のパラメータのデフォルト値
const _DEFAULT_MODEL = "gpt-4"
const _DEFAULT_API_VERSION = "2024-02-15-preview" // Azure用
const _DEFAULT_MAX_TOKENS = 2000;
const _DEFAULT_TEMPERATURE = 0.6;
const _DEFAULT_MAX_RETRY = 5;
const _DEFAULT_MAX_RETRY_FOR_FORMAT_AI_MESSAGE = 5;

// 画像生成AI向けのデフォルトパラメータ
//
// 本家OpenAI
// https://platform.openai.com/docs/api-reference/images/create
// 
// Azure
// https://learn.microsoft.com/ja-jp/azure/ai-services/openai/reference#request-a-generated-image-dall-e-3
const _DEFAULT_IMAGE_N = 1;
const _DEFAULT_IMAGE_SIZE = "1024x1024";
const _DEFAULT_IMAGE_QUALITY = "standard";
const _DEFAULT_IMAGE_RESPONSE_FORMAT = "url";
const _DEFAULT_IMAGE_STYLE = "vivid";


/**
 * AzureOpenAI オブジェクトを生成します。AzureのOpenAI APIとの通信を管理するためのクライアントです。
 * このオブジェクトを通じて、Azure上でホストされているOpenAIのモデル（例えば、GPT-4）を利用して
 * テキスト生成やその他のAIに関連するタスクを実行することができます。
 *
 * 主な機能としては、プロンプトをAIに送信し、生成されたテキストやJSONスキーマに基づく
 * レスポンスを取得することができます。また、APIキー、モデルの種類、トークンの最大数、
 * 温度パラメータ、最大リトライ回数など、APIリクエストに関連する複数の設定をカスタマイズ可能です。
 *
 * 使用方法:
 * const client = OpenAI.createAzureOpenAIClient({
 *   apiKey: '<YOUR_API_KEY>',
 *   azureEndpoint: '<AZURE_ENDPOINT>',
 * });
 * const response = client.simpleChat("Hello, world!");
 * 
 * 
 * @param {Object} config - オブジェクトの構成オプションを含む設定オブジェクト。
 * @param {string} config.apiKey - APIキーの文字列。必須です。
 * @param {string} config.azureEndpoint - Azure エンドポイントURLです。
 * @param {string} [config.model="gpt-4"] - 使用するモデルの識別子。省略可能で、デフォルトは gpt-4 です。
 * @param {number} [config.maxTokens=2000] - トークンの最大数。省略可能で、デフォルトは 2000 です。
 * @param {number} [config.temperature=0.6] - モデルの温度パラメータ。省略可能で、デフォルトは 0.6 です。
 * @param {Blob[]} [config.images] - 画像です。gpt-4-vision系モデルじゃないと扱えないので、これを指定するときは model パラメータも適切に指定してください。
 * @param {Object} [config.responseSchema] - AIからの出力フォーマットを表すJSONスキーマ。
 * @param {Object[]} [config.functions] - AIが必要に応じて実行する関数のオプションのリスト。各オブジェクトは以下のプロパティを持つ:
 * @param {Function} config.functions[].func - 必要に応じて実行する関数。
 * @param {string} config.functions[].description - 関数の説明。
 * @param {Object} config.functions[].parameters - 関数の引数を定義するJSONスキーマ。
 * @param {Function[]} [config.functions] - AIに必要に応じて実行させる関数のリスト。
 * @param {number} [config.maxRetry=5] - 最大リトライ回数。省略可能で、デフォルトは5です。
 * @param {number} [config.maxRetryForFormatAiMessage=5] - responseSchema指定時にレスポンスJSON化の最大リトライ回数。省略可能で、デフォルトは5です。
 */
function createAzureOpenAIClient(config) {
  return new AzureOpenAI(config);
}

/**
 * 本家OpenAI の OpenAI API にアクセスするクライアントオブジェクトを生成します。
 * Azure へのアクセスはこちらではなく、createAzureOpenAIClient() を使用してください。
 * 
 * 使用方法:
 * const client = OpenAI.createOpenAIClient({
 *   apiKey: '<YOUR_API_KEY>',
 * });
 * const response = client.simpleChat("What is the meaning of life?");
 * 
 * 
 * @param {Object} config - オブジェクトの構成オプションを含む設定オブジェクト。
 * @param {string} config.apiKey - APIキーの文字列。必須です。
 * @param {string} [config.model="gpt-4"] - 使用するモデルの識別子。省略可能で、デフォルトは gpt-4 です。
 * @param {number} [config.maxTokens=2000] - トークンの最大数。省略可能で、デフォルトは 2000 です。
 * @param {number} [config.temperature=0.6] - モデルの温度パラメータ。省略可能で、デフォルトは 0.6 です。
 * @param {Blob[]} [config.images] - 画像です。gpt-4-vision系モデルじゃないと扱えないので、これを指定するときは model パラメータも適切に指定してください。
 * @param {Object} [config.responseSchema] - AIからの出力フォーマットを表すJSONスキーマ。
 * @param {Object[]} [config.functions] - AIが必要に応じて実行する関数のオプションのリスト。各オブジェクトは以下のプロパティを持つ:
 * @param {Function} config.functions[].func - 必要に応じて実行する関数。
 * @param {string} config.functions[].description - 関数の説明。
 * @param {Object} config.functions[].parameters - 関数の引数を定義するJSONスキーマ。
 * @param {number} [config.maxRetry=5] - 最大リトライ回数。省略可能で、デフォルトは5です。
 * @param {number} [config.maxRetryForFormatAiMessage=5] - responseSchema指定時にレスポンスJSON化の最大リトライ回数。省略可能で、デフォル */
function createOpenAIClient(config) {
  return new OpenAI(config);
}

class BaseOpenAI {
  // TODO パラメータのバリデーションをする。そうじゃないと間違ったパラメータを与えてしまったときに,なぜ動作に反映されないのか気づくのが難しい
  constructor(config) {
    this.apiKey = config.apiKey;
    this.model = config.model || _DEFAULT_MODEL;
    this.maxTokens = config.maxTokens || _DEFAULT_MAX_TOKENS;
    this.temperature = config.temperature || _DEFAULT_TEMPERATURE;
    this.responseSchema = config.responseSchema;
    this.maxRetry = config.maxRetry || _DEFAULT_MAX_RETRY;
    this.maxRetryForFormatAiMessage = config.maxRetryForFormatAiMessage || _DEFAULT_MAX_RETRY_FOR_FORMAT_AI_MESSAGE;

    const requiredFields = ['apiKey'];
    requiredFields.forEach(x => {
      if (!this[x]) {
        throw new Error(x+'が指定されていません。');
      }
    });
  }

  /**
   * AIにプロンプトを渡して文字列を生成させます。
   * params では今回の呼び出しにのみ適用されるパラメータを指定可能です。
   * 省略するとインスタンス化時に設定した値になります。
   * 
   * @param {Object} [params] - 生成オプションを含む設定オブジェクト。
   * @param {string} [params.model] - 使用するモデルの識別子。
   * @param {number} [params.maxTokens] - トークンの最大数。
   * @param {number} [params.temperature] - モデルの温度パラメータ。
   * @param {Blob[]} [params.images] - 画像です。gpt-4-vision系モデルじゃないと扱えないので、これを指定するときは model パラメータも適切に指定してください。
   * @param {Object} [params.responseSchema] - AIからの出力フォーマットを表すJSONスキーマ。
   * @param {Object[]} [params.functions] - AIが必要に応じて実行する関数のオプションのリスト。各オブジェクトは以下のプロパティを持つ:
   * @param {Function} params.functions[].func - 必要に応じて実行する関数。
   * @param {string} params.functions[].description - 関数の説明。
   * @param {Object} params.functions[].parameters - 関数の引数を定義するJSONスキーマ。
   * @param {number} [params.maxRetry] - 最大リトライ回数。
   * @param {number} [params.maxRetryForFormatAiMessage=5] - responseSchema指定時にレスポンスJSON化の最大リトライ回数。省略可能で、デフォル
   * @return {Object|string} prams.responseSchema を指定していればその型のオブジェクト、そうでなければテキスト
   * @throws {Error} OpenAI APIレイヤでのエラーが発生した場合に例外をスローします。
   */
  simpleChat(prompt, params={}) {
    const result = this.chatCompletion(prompt, params);
    if (result.error) {
      throw new Error("API エラー: " + JSON.stringify(result.error));
    }

    if (result.choices[0].message.function_call) {
      let resObj = null;
      const argJson = result.choices[0].message.function_call.arguments;

      try {
        resObj = JSON.parse(argJson);
      } catch (e) {
        throw new Error("JSONパースに失敗しました。完全なJSONになっていない場合、params.maxTokensを増やしてみてください。: argJson=" + argJson + ", 元のError=" + e.toString());
      }

      return resObj;
    }

    return result.choices[0].message.content;
  }

  /**
   * AIにプロンプトを渡して文字列を生成させます。
   * params では今回の呼び出しにのみ適用されるパラメータを指定可能です。
   * 省略するとインスタンス化時に設定した値になります。
   * 
   * @param {Object} [params] - 生成オプションを含む設定オブジェクト。
   * @param {string} [params.model] - 使用するモデルの識別子。
   * @param {number} [params.maxTokens] - トークンの最大数。
   * @param {number} [params.temperature] - モデルの温度パラメータ。
   * @param {Object} [params.responseSchema] - AIからの出力フォーマットを表すJSONスキーマ。
   * @param {Object[]} [params.functions] - AIが必要に応じて実行する関数のオプションのリスト。各オブジェクトは以下のプロパティを持つ:
   * @param {Function} params.functions[].func - 必要に応じて実行する関数。
   * @param {string} params.functions[].description - 関数の説明。
   * @param {Object} params.functions[].parameters - 関数の引数を定義するJSONスキーマ。
   * @param {Blob[]} [params.images] - 画像です。gpt-4-vision系モデルじゃないと扱えないので、これを指定するときは model パラメータも適切に指定してください。
   * @param {number} [params.maxRetry] - 最大リトライ回数。
   * @param {number} [params.maxRetryForFormatAiMessage=5] - responseSchema指定時にレスポンスJSON化の最大リトライ回数。省略可能で、デフォル   
   * @return {Object} OpenAI APIからのレスポンスJSONをパースしたオブジェクト
   */
  chatCompletion(prompt, params={}) {
    const payload = {
      messages: [
        {
          role: "user",
          content: prompt
        }
      ]
    };

    // 都度上書き可能なパラメーター
    ["model", "max_tokens", "temperature", ].forEach(x => {
      const key = this.toCamelCase_(x);
      const value = params[key] || this[key];

      if (value) {
        payload[x] = value;
      }
    });

    // JSONスキーマが指定されていれば Function Calling を強制して、JSONでレスポンスを得る
    const responseSchema =  params.responseSchema || this.responseSchema;

    if (responseSchema) {
      payload.functions = [
        {
          name: "format_ai_message",
          description: "最終的なAIの応答をフォーマットするための関数です。最後に必ず実行してください。",
          parameters: responseSchema
        }
      ]; 
      payload.function_call = {name: "format_ai_message"}
    }

    if (params.functions) {
      if (payload.functions == null) {
        payload.functions = [];
      }

      params.functions.forEach(f => {
        payload.functions.push({
          name: f.func.name,
          description: f.description,
          parameters: f.parameters
        });

        payload.function_call = "auto";
      });
    }


    // 画像が指定されていれば画像もGPTへ渡す。gpt-4-vision系モデルじゃないとエラーになる
    const images =  params.images || this.images;

    if (images) {
      // contentを配列化
      payload.messages[0].content = [{
        "type": "text",
        "text": payload.messages[0].content
      }];      

      images.forEach(imageBlob => {
        const mime_type = imageBlob.getContentType();
        const image_bin = imageBlob.getBytes();
        const image_b64 = Utilities.base64Encode(image_bin);
        const dataUriScheme = `data:${mime_type};base64,${image_b64}`;

        payload.messages[0].content.push({
          'type': 'image_url',
          'image_url': {
            'url': dataUriScheme
          }
        });
      });
    }

    const url = this.getChatCompletionUrl_(params);

    let retryForFormatAiMessage = 0;
    const maxRetryForFormatAiMessage = params.maxRetryForFormatAiMessage || this.maxRetryForFormatAiMessage;

    // Function Callingの処理
    // AI が Function を要求してきていれば呼び出して、結果をAIに伝える
    while (true) {
      const res = this.callApi_(url, payload, params.maxRetry);

      if (res.error != null) {
        return res;
      }

      if (!res.choices[0].message.function_call) {
        if (responseSchema) {
          // 必要なFunctionが呼び出されていない場合は再試行。
          if (retryForFormatAiMessage < maxRetryForFormatAiMessage) {
            Logger.log("format_ai_message function is not called. retrying..: retryCont=" + retryForFormatAiMessage);
            retryForFormatAiMessage++;
            continue;
          }

          throw new Error("format_ai_message function のリトライ最大回数に到達しましたが、実行されませんでした。");
        } else {
          // functionsの指定もなく、responseSchemaの指定もないシンプルな応答
          return res;
        }
      }

      if (res.choices[0].message.function_call.name == "format_ai_message") {
        // これ以上のFunction callの必要がなければ終了
        return res;
      }

      const functionCall = res.choices[0].message.function_call;
      const targetFunction = params.functions.find(x => x.func.name == functionCall.name).func;
      const funcArgs = JSON.parse(functionCall.arguments);

      const funcResult = targetFunction(funcArgs);
      const funcResultJson = JSON.stringify(funcResult);

      // payload.messages.push(res.choices[0].message);
      payload.messages.push({"role": "function", "name": functionCall.name, "content": funcResultJson});
    }

  }
  /**
   * 音声の文字起こしを行います。
   * 文字起こしされたテキストのみを返します。
   * Azure の場合は25MiB以下のWavファイルのみ(だと思う:-p)。
   * 
   * @param {Blob} audio 音声ファイルのBlob
   * @return {string} 文字起こしされたテキスト
   */

  simpleVoiceToText(audio, params={}) {
    return this.audioTranscription(audio, params).text;
  }

  /**
   * 音声の文字起こしを行います。
   * Azure の場合は25MiB以下のWavファイルのみ(だと思う:-p)。
   * 
   * @param {Blob} audio 音声ファイルのBlob
   * @return {Object} OpenAI APIからのレスポンスJSONをパースしたオブジェクト
   */
  audioTranscription(audio, params={}) {

    // TODO 大きなファイルの分割処理や、非対応フォーマットから対応フォーマットへの変換などなどへの対応。

    const url = this.getAudioTranscriptionUrl_(params);

    const payload = {
      model: params.model || this.model,
      file: audio
    };

    return this.callApiMultipart_(url, payload, params.maxRetry);
  }

  /**
   * 画像を生成します。
   * 詳しいパラメータは以下を参照。
   *
   * 本家OpenAI
   * https://platform.openai.com/docs/api-reference/images/create
   * 
   * Azure
   * https://learn.microsoft.com/ja-jp/azure/ai-services/openai/reference#request-a-generated-image-dall-e-3
   * 
   * @param {string} prompt - プロンプト
   * @param {Object} [params] - 生成オプションを含む設定オブジェクト。
   * @param {string} [params.model] - 使用するモデルの識別子。
   * @param {number} [params.n] - 生成する画像の枚数。
   * @param {number} [params.size] - 画像サイズ 。
   * @param {number} [params.quality] - 画質。
   * @param {number} [params.response_format] - レスポンスフォーマット。
   * @param {number} [params.style] - スタイル。
   * @return {Object} response_format=rulのときは、url, response_format=b64_jsonのときはBase64エンコードされた画像
   */
  simpleImageGeneration(prompt, params={}) {
    const result = this.imageGeneration(prompt, params);

    return result.data[0].url || result.data[0].b64_json
  }

  /**
   * 画像を生成します。
   * 詳しいパラメータは以下を参照。
   *
   * 本家OpenAI
   * https://platform.openai.com/docs/api-reference/images/create
   * 
   * Azure
   * https://learn.microsoft.com/ja-jp/azure/ai-services/openai/reference#request-a-generated-image-dall-e-3
   * 
   * @param {string} prompt - プロンプト
   * @param {Object} [params] - 生成オプションを含む設定オブジェクト。
   * @param {string} [params.model] - 使用するモデルの識別子。
   * @param {number} [params.n] - 生成する画像の枚数。
   * @param {number} [params.size] - 画像サイズ 。
   * @param {number} [params.quality] - 画質。
   * @param {number} [params.response_format] - レスポンスフォーマット。
   * @param {number} [params.style] - スタイル。
   * @return {Object} OpenAI APIからのレスポンスJSONをパースしたオブジェクト
   */
  imageGeneration(prompt, params={}) {
    const payload = {
      "prompt": prompt,
      "n": params.n || _DEFAULT_IMAGE_N,
      "size": params.size || _DEFAULT_IMAGE_SIZE,
      "quality": params.quality || _DEFAULT_IMAGE_QUALITY,
      "response_format": params.response_format || _DEFAULT_IMAGE_RESPONSE_FORMAT,
      "style": params.style || _DEFAULT_IMAGE_STYLE,
    };

    const url = this.getImageGenerationUrl_(params);

    return this.callApi_(url, payload, params.maxRetry);
  }

  /**
   * Web APIをコールします。
   * Content-Type が application/json のリクエストを行います。
   * HTTPステータス429(Too Many Requests)時に Retry-After ヘッダーで何秒後にアクセスすれば良いか指示された場合だけリトライします。
   *
   * @param {string} url - APIエンドポイントのURL。
   * @param {Object} payload - ペイロード。
   * @param {number} [maxRetry=this.maxRetry] - 最大リトライ回数。
   */
  callApi_(url, payload, maxRetry=this.maxRetry) {
    Logger.log(`accessing url: ${url}`);
    Logger.log("payload: " + JSON.stringify(payload));
    // Logger.log("params: " + JSON.stringify(params));

    const headers = this.getAuthorizationHeader_();

    const options = {
      method: 'post',
      contentType: 'application/json',
      headers: headers,
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };

    return this.requestWithRetry_(url, options, maxRetry);
  }

  /**
   * Web APIをコールします。
   * Content-Type が mulipart/form-data のリクエストを行います。
   * HTTPステータス429(Too Many Requests)時に Retry-After ヘッダーで何秒後にアクセスすれば良いか指示された場合だけリトライします。
   *
   * @param {string} url - APIエンドポイントのURL。
   * @param {Object} payload - ペイロード。
   * @param {number} [maxRetry=this.maxRetry] - 最大リトライ回数。
   */
  callApiMultipart_(url, payload, maxRetry=this.maxRetry) {
    const headers = this.getAuthorizationHeader_();

    var options = {
      method: "post",
      headers: headers,
      payload: payload,
      muteHttpExceptions: true
    };

    return this.requestWithRetry_(url, options, maxRetry);
  }

  // リトライ制御しつつHTTPリクエストするメソッド
  requestWithRetry_(url, options, maxRetry=this.maxRetry) {
    for (let attempts = 0; attempts < maxRetry; attempts++) {
      try {
        const response = UrlFetchApp.fetch(url, options);
        // Logger.log('Headers: ' + JSON.stringify(response.getHeaders()));

        const content = response.getContentText();
        Logger.log('contentText: ' + content);

        // リトライ条件
        const httpStatus = response.getResponseCode();
        if (httpStatus == 429) {
          const headers = response.getHeaders();
          const retryAfter = this.getDictValue_('retry-after', headers) || this.getDictValue_('x-ratelimit-reset-tokens', headers);
          if (retryAfter) {
            Logger.log(`Rate limit exeeds. Retry after ${retryAfter} seconds.`);
            Utilities.sleep(retryAfter * 1000);
            continue;
          }
        } else if (httpStatus != 200) {
          throw new Error(`APIエラー: status=${httpStatus}, message=${content}`);
        }

        const json = JSON.parse(content);
        return json;
      } catch (e) {
        const message = e.toString();
        Logger.log(message);
        if (attempts >= 2 || response.getResponseCode() !== 429) {
          throw new Error(message);
        }
      }
    }

    // リトライが最大回数に達した
    throw new Error(`APIエラー: リトライが最大回数に達しました： maxRetry=${this.maxRetry}`);
  }

  /**
   * 文字列をキャメルケースに変換します。
   * @param {string} str 文字列
   * @return {string} キャメルケースの文字列
   */
  toCamelCase_(str) {
    str = str.charAt(0).toLowerCase() + str.slice(1);
    return str.replace(/[-_](.)/g, function(match, group1) {
        return group1.toUpperCase();
    });
  }

  /**
   * 指定されたキーに対応する連想配列（オブジェクト）の値を大文字小文字を無視して取得します。
   * キーが存在しない場合、undefinedが返されます。
   * @param {string} key 検索するキー
   * @param {Object} dict 検索対象の連想配列（オブジェクト）
   * @return {any} 指定されたキーに対応する値、またはキーが存在しない場合はundefined
   */
  getDictValue_(key, dict) {
    const normalizedKey = key.toLowerCase();
    const foundKey = Object.keys(dict).find(dictKey => dictKey.toLowerCase() === normalizedKey);
    return foundKey ? dict[foundKey] : undefined;
  }

  // 抽象メソッド
  getChatCompletionUrl_(params={}) {}

  // 抽象メソッド
  getAudioTranscriptionUrl_(params={}) {}

  // 抽象メソッドオーバーライド
  getImageGenerationUrl_(params={}) {}

  // 抽象メソッド
  getAuthorizationHeader_(){}
}

/**
 * AzureOpenAI クラスは、AzureのOpenAI APIとの通信を管理するためのクライアントです。
 * このクラスを通じて、Azure上でホストされているOpenAIのモデル（例えば、GPT-4）を利用して
 * テキスト生成やその他のAIに関連するタスクを実行することができます。
 *
 * 主な機能としては、プロンプトをAIに送信し、生成されたテキストやJSONスキーマに基づく
 * レスポンスを取得することができます。また、APIキー、モデルの種類、トークンの最大数、
 * 温度パラメータ、最大リトライ回数など、APIリクエストに関連する複数の設定をカスタマイズ可能です。
 * 
 * 使用方法:
 * const client = new AzureOpenAIClient({
 *   apiKey: '<YOUR_API_KEY>',
 *   azureEndpoint: '<AZURE_ENDPOINT>',
 * });
 * const response = client.simpleChat("Hello, world!");
 * 
 */
class AzureOpenAI extends BaseOpenAI {

  constructor(config) {
    super(config);
    this.azureEndpoint = config.azureEndpoint;
    this.apiVersion = config.apiVersion || _DEFAULT_API_VERSION;

    const requiredFields = ['azureEndpoint'];
    requiredFields.forEach(x => {
      if (!this[x]) {
        throw new Error(x+'が指定されていません。');
      }
    });
  }

  // 抽象メソッドオーバーライド
  getChatCompletionUrl_(params={}) {
    const url = this.azureEndpoint + "/openai/deployments/" + (params.model || this.model) + "/chat/completions?api-version=" + this.apiVersion;
    return url;
  }

  // 抽象メソッドオーバーライド
  getAudioTranscriptionUrl_(params={}) {
    const url = this.azureEndpoint + "/openai/deployments/" + (params.model || this.model) + "/audio/transcriptions?api-version=" + this.apiVersion;
    return url;
  }

  // 抽象メソッドオーバーライド
  getImageGenerationUrl_(params={}) {
    const url = this.azureEndpoint + "/openai/deployments/" + (params.model || this.model) + "/images/generations?api-version=" + this.apiVersion;
    return url;
  }

  // 抽象メソッドオーバーライド
  getAuthorizationHeader_(){
    return {"api-key": this.apiKey}
  }
}

/**
 * OpenAI クラスは、本家OpenAIのAPIとの通信を管理するためのクライアントです。
 *
 * 使用方法:
 * const client = new OpenAI({
 *   apiKey: '<YOUR_API_KEY>',
 * });
 * const response = client.simpleChat("What is the meaning of life?");
 */
class OpenAI extends BaseOpenAI {

  constructor(config) {
    super(config);
  }

  // 抽象メソッドオーバーライド
  getChatCompletionUrl_(params={}) {
    return "https://api.openai.com/v1/chat/completions";
  }

  // 抽象メソッドオーバーライド
  getAudioTranscriptionUrl_(params={}) {
    const url = "https://api.openai.com/v1/audio/transcriptions";
    return url;
  }

  // 抽象メソッドオーバーライド
  getImageGenerationUrl_(params={}) {
    const url = "https://api.openai.com/v1/images/generations";
    return url;
  }

  // 抽象メソッドオーバーラド
  getAuthorizationHeader_(){
    return {'Authorization': 'Bearer ' + this.apiKey};
  }
}
