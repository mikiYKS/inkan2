  var dt = new Date();
  var txtDate = dt.getFullYear().toString() + "-" + (dt.getMonth() + 1) + "-" + dt.getDate();
  var authenticator;
  var client_id = "f00c6849-1c45-474e-8aa4-4b7bd80dd530";
  var redirect_url = "https://mikiyks.github.io/inkan2/";
  var scope;
  var access_token;

Office.onReady(function () {
  getUser();
  $("#date").val(txtDate);
  $("#run").click(() => tryCatch(run));
  //日付不要にチェック入れたら日付グレーアウト
  $("#dateCheckBox").change(() => tryCatch(change));
  function change() {
    if ($("#dateCheckBox").prop("checked")) {
      $("#date").prop("disabled", true);
    } else {
      $("#date").prop("disabled", false);
    }
  }
});

async function run() {
  await Word.run(async (context) => {
    //名前が空なら処理なし
    if (
      !$("#name").val()
    ) {
    } else {
      //印鑑生成実行
      inkanOnCanvas();
      //ログ出力
      $(function () {
        Office.context.document.getFilePropertiesAsync(async function (asyncResult) {
          var fileUrl = asyncResult.value.url;
          var fileName;
          var inkanName = $("#name").val();
          if (fileUrl == "") {
            fileName = '未保存ワード';
          } else {
            fileName = fileUrl.match(".+/(.+?)([\?#;].*)?$")[1];
          };
          inkanLog(inkanName, fileName);
        });
      });
    }
  });
}

async function tryCatch(callback) {
  try {
    await callback();
  } catch (error) {
    console.error(error);
  }
}

//キャンバスに印影作成
async function inkanOnCanvas() {
  var canvas = document.querySelector("#canvas");
  var ctx = canvas.getContext("2d");
  var nametxt = $("#name")
    .val()
    .toString()
    .replace("（", "")
    .replace("(", "")
    .replace("）", "")
    .replace(")", "")
    .replace(" ", "")
    .replace("　", "");
  var lenname = nametxt.length;
  var fsize = 55 - (7 / 3) * (lenname - 1);
  var dateText =
    "'" +
    $("#date")
      .val()
      .toString()
      .slice(2, 4) +
    "." +
    $("#date")
      .val()
      .toString()
      .slice(5, 7) +
    "." +
    $("#date")
      .val()
      .toString()
      .slice(8, 10);

  ctx.setTransform(1, 0, 0, 1, 0, 0);
  ctx.font = fsize + "pt HGS行書体, HGS明朝E";
  var namewidth = ctx.measureText(nametxt).width;

  ctx.setTransform(1, 0, 0, 80 / namewidth, 0, (80 / namewidth) * 65.332 + 6.5335);
  ctx.font = fsize + "pt HGS行書体, HGS明朝E";
  ctx.fillStyle = "rgba(255, 32, 0)";
  tategaki(ctx, nametxt, 0);

  ctx.setTransform(1, 0, 0, 1, 0, 0);
  ctx.beginPath();
  ctx.arc(50, 50, 45, 0, Math.PI * 2, true);
  ctx.strokeStyle = "rgba(255, 32, 0)";
  ctx.lineWidth = 2;
  ctx.stroke();

  if ($("#dateCheckBox").prop("checked")) {
  } else {
    ctx.setTransform(1, 0, 0, 1, 0, 0);
    ctx.font = 16 + "pt Calibri bold";
    ctx.fillStyle = "rgba(255, 32, 0)";
    ctx.fillText(dateText, 50 - ctx.measureText(dateText).width / 2, 118);
  };

  var nameBase64Img = canvas.toDataURL().replace(/^.*,/, "");

  insertImage(nameBase64Img);

  ctx.clearRect(0, 0, 100, 120);
}

//カーソル位置に印影貼り付け
async function insertImage(base64img) {
  await Word.run(async (context) => {
    context.document.getSelection().insertInlinePictureFromBase64(base64img, "End").height = 43;
    await context.sync();
  });
}

//縦書き変換
function tategaki(context, text, y) {
  var textList = text.split("\n");
  var lineHeight = context.measureText("あ").width;
  textList.forEach(function (elm, i) {
    Array.prototype.forEach.call(elm, function (ch, j) {
      context.fillText(ch, 50 - lineHeight / 2, y + lineHeight * j);
    });
  });
}


Office.initialize = function (reason) {
  if (OfficeHelpers.Authenticator.isAuthDialog()) return;
};

async function getUser() {
  scope = "https://graph.microsoft.com/user.read";
  authenticator = new OfficeHelpers.Authenticator();
  //access_token取得
  authenticator.endpoints.registerMicrosoftAuth(client_id, {
    redirectUrl: redirect_url,
    scope: scope
  });
  //認証
  authenticator
    .authenticate(OfficeHelpers.DefaultEndpoints.Microsoft)
    .then(function (token) {
      access_token = token.access_token;
      $("#exec").prop("disabled", false);
      //API呼び出し　ユーザー情報取得
      $(function () {
        $.ajax({
          url: "https://graph.microsoft.com/v1.0/me",
          type: "GET",
          beforeSend: function (xhr) {
            xhr.setRequestHeader("Authorization", "Bearer " + access_token);
          },
          success: function (data) {
            //取得した苗字をセット
            $("#name").val(data.surname);
          },
          error: function (data) {
            console.log(data);
          }
        });
      });
    })
    .catch(OfficeHelpers.Utilities.log);
}

//SharePointListにログ出力
function inkanLog(inkanName, inkanFile) {
  scope = "https://graph.microsoft.com/Sites.ReadWrite.All";
  authenticator = new OfficeHelpers.Authenticator();
  //access_token取得
  authenticator.endpoints.registerMicrosoftAuth(client_id, {
    redirectUrl: redirect_url,
    scope: scope
  });
  //認証
  authenticator.authenticate(OfficeHelpers.DefaultEndpoints.Microsoft).then(function (token) {
    access_token = token.access_token;
    //API呼び出し印鑑ログ投稿
    $(function () {
      $.ajax({
        url:
          "https://graph.microsoft.com/v1.0/sites/20531fc2-c6ab-4e1e-a532-9c8e15afed0d/lists/6aac0560-622e-4ee1-ba8f-73b32d8e9f05/items",
        type: "POST",
        data: JSON.stringify({
          fields: {
            Title: inkanName,
            FileName: inkanFile
          }
        }),
        contentType: "application/json",
        beforeSend: function (xhr) {
          xhr.setRequestHeader("Authorization", "Bearer " + access_token);
        }
      }).then(
        async function (data) { },
        function (data) {
          console.log(data);
        }
      );
    });
  });
}