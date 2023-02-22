const proxy = "http://localhost:5000/proxy"; // https://api.openai.com/v1/completions
const prompts = [
  `I want you to act as an Email subject generator. I will give you email text in english language and you will understand it and then suggest me a single concise and professional subject for my email. I want you to limit the words to 10 for each subject. My email is, \n`,
  "I want to act like a AI assisted email reader. I will give you an email body and you will read it and extract important information or action-items which needs my attention and then list them in concise and unique points not more than 10 words in each point. the number of points should be as less as possible. I don't expect any heading, just the direct point, I want. I want in english. the email is:\n",
  "I want you to act as an Email formatter, spelling corrector and improver. I will give you an email in english and you will give me improved version of my mail, in English. Keep the meaning same, but make the mail more literary and professional. My email is:\n"
];

export async function callLLMApi(
  query = "Seattle is",
  modelName = "text-davinci-003",
  count = 5,
  best_of = 10,
  max_tokens = 100,
  temperature = 0.7,
  pIdx = 0
) {
  const myHeaders = new Headers();
  myHeaders.set("X-ModelType", modelName);

  const raw = JSON.stringify({
    prompt: prompts[pIdx] + query,
    max_tokens: max_tokens, // for summary => 500
    temperature: temperature, // for summary=> temp:1.4
    top_p: 1,
    n: count, // for summary => count:1
    best_of: best_of,
    stream: false,
    logprobs: null,
    //"stop": "\n"
  });

  var requestOptions = {
    method: "POST",
    headers: myHeaders,
    body: raw,
    redirect: "follow",
    // mode: 'no-cors'
  };

  return await fetch(proxy, requestOptions)
    .then(async (response) => {
      if (response.ok) {
        const data = await response.json();
        return data.choices;
      }
      console.log("status not ok: ", response.statusText);
    })
    .then((data) => {
      // console.log(data[0].text);
      data.forEach((e) => {
        console.log(e.text);
      });
      return data;
    })
    .catch((error) => console.log("error", error));
}
