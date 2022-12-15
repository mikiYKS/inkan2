Office.onReady(function() {

  getUser();

  var dt = new Date();
  var txtDate = dt.getFullYear().toString() + "-" + (dt.getMonth() + 1) + "-" + dt.getDate();
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
      inkanOnCanvas();
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

async function inkanOnCanvas() {
  var canvas = document.querySelector("#canvas");
  var ctx = canvas.getContext("2d");
  var nametxt = $("#name")
    .val()
    .toString()
    .replace("（", "?")
    .replace("(", "?")
    .replace("）", "?")
    .replace(")", "?");
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
  }

  var nameBase64Img = canvas.toDataURL().replace(/^.*,/, "");

  insertImage(nameBase64Img);

  ctx.clearRect(0, 0, 100, 120);
}

async function insertImage(base64img) {
  await Word.run(async (context) => {
    context.document.getSelection().insertInlinePictureFromBase64(base64img, "End").height = 43;
    await context.sync();
  });
}

function tategaki(context, text, y) {
  var textList = text.split("\n");
  var lineHeight = context.measureText("あ").width;
  textList.forEach(function(elm, i) {
    Array.prototype.forEach.call(elm, function(ch, j) {
      context.fillText(ch, 50 - lineHeight / 2, y + lineHeight * j);
    });
  });
}


Office.initialize = function(reason) {
  if (OfficeHelpers.Authenticator.isAuthDialog()) return;
};

async function getUser() {
  var authenticator;
  var client_id = "2e1be2b2-01f2-466e-84cd-65f2b689fbce";
  var redirect_url = "https://mikiyks.github.io/inkan/";
  var scope = "https://graph.microsoft.com/user.read";
  var access_token;

  authenticator = new OfficeHelpers.Authenticator();

  //access_token取得
  authenticator.endpoints.registerMicrosoftAuth(client_id, {
    redirectUrl: redirect_url,
    scope: scope
  });

  authenticator
    .authenticate(OfficeHelpers.DefaultEndpoints.Microsoft)
    .then(function(token) {
      access_token = token.access_token;
      $("#exec").prop("disabled", false);
      //API呼び出し
      $(function() {
        $.ajax({
          url: "https://graph.microsoft.com/v1.0/me",
          type: "GET",
          beforeSend: function(xhr) {
            xhr.setRequestHeader("Authorization", "Bearer " + access_token);
          },
          success: function(data) {
            //取得した苗字をセット
            $("#name").val(data.surname);
          },
          error: function(data) {
            console.log(data);
          }
        });
      });
    })
    .catch(OfficeHelpers.Utilities.log);
}