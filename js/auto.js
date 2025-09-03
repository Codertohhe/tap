// Firefox 1.0+
var isFirefox = typeof InstallTrigger !== "undefined";

var preservemyHTML = "";

function addToTestAuto(selectedQuestions) {
  resetSummary();

  testBank = [];
  for (var i = 0; i < selectedQuestions.length; i++) {
    testBank.push(String(selectedQuestions[i]));
  }

  $("#typeAuto #pdfPreview").empty();

  $userInfo = $('<div id="userInfo"></div>');
  $userInfo.append(
    '<div id="form-shcool"><label>School:</label> <input type="text" name="school" id="school" style="margin-top: 10px;margin-left: 10px;" /></div>'
  );
  $userInfo.append(
    '<div id="form-subject"><label>Subject:</label> <input type="text" name="subject" value="' +
      allQuestions().first().subject +
      '" id="subject"  /></div>'
  );
  $userInfo.append(
    '<div id="form-class"><label>Class:</label> <input type="text" name="class_t" value="' +
      '" id="class_t" style="margin-left: 12px;"/></div>'
  );
  //$userInfo.append('<div id="form-roll"><label>Roll no:</label> <input type="text" name="roll" id="roll" style="margin-left: 1px;margin-top: 0px;}" /></div>');
  $userInfo.append(
    '<div id="form-name"><label>Name:</label> <input type="text" name="name" id="name" style="margin-left: 14px;" /></div>'
  );
  $userInfo.append(
    '<div id="form-time"><label>Time:</label> <input type="text" name="time" onkeydown="return numValidate(event)" id="time" style="margin-left: 12px; margin-top: -3px;" /></div>'
  );
  //$userInfo.append('<div id="form-set"><label>Set no:</label> <input type="text" name="set" id="set" style="margin-left: 7px;" /></div>');
  $userInfo.append(
    '<div id="logoUpdateParent" style="width: 345px;"><label>Logo: </label> <input type="file" accept="image/*" name="imgfile" id="imgfile" size="50"/><div class="fileChooseBlocker"></div></div>'
  );
  $userInfo.append(
    '<div id="form-btn"><button id="showPrevBtn" class="sm_button" onclick="showPreviewAuto()"  >Submit</button></div>'
  );
  $userInfo.append(
    '<button id="closepopupBtn" class="" title="Close" onclick="closePopupFromPreview()"><i class="fa fa-times"></i></i></button>'
  );

  $userInfo.append('<div class="clear"></div>');

  $("#typeAuto #pdfPreview").append($userInfo);
  testBank = uniqueArray(testBank);
  console.log(testBank);

  selectChapters = [];

  for (var i = 0; i < testBank.length; i++)
    insertQuestion($("#typeAuto #pdfPreview"), testBank[i]);

  for (var zz = 0; zz < selectChapters.length; zz++) {
    $("#" + selectChapters[zz]).addClass("practive");
  }

  var categoriges = $("#typeAuto #pdfPreview #sContainIn")
    .children()
    .map(function () {
      return this.dataset.title;
    })
    .get();
  var chapters = $("#typeAuto #pdfPreview #sContainIn")
    .children()
    .map(function () {
      return this.dataset.chapter;
    })
    .get();
  var type = $("#pdfPreview #sContainIn")
    .children()
    .map(function () {
      return this.dataset.type;
    })
    .get();

  var marks = $("#pdfPreview #sContainIn")
    .children()
    .map(function () {
      return this.dataset.marks;
    })
    .get();

  categoriges = uniqueArray(categoriges);
  chapters = uniqueArray(chapters);

  /*console.log(chapters);
	console.log(categoriges);*/

  var uniquetype = [];
  $.each(type, function (i, el) {
    if ($.inArray(el, uniquetype) === -1) uniquetype.push(el);
  });

  /* 	For filtering and merging questions by type */

  var $section;
  var totalCat_marks;

  var tempinc = 1;

  filterArr = [];

  for (var catIndex = 0; catIndex < uniquetype.length; catIndex++) {
    var obj = new Object();

    obj.type = uniquetype[catIndex];
    obj.arr = [];
    obj.Finalarr = [];
    obj.Title = [];
    obj.FinalTitle = [];
    obj.ID = [];
    obj.outcome = [];
    filterArr.push(obj);

    tempinc = 1;
  }

  $("#sContainIn .question-parent").each(function (i, e) {
    for (var o in filterArr) {
      if ($(this).attr("data-type") == filterArr[o].type) {
        filterArr[o].arr.push(this.dataset.type);
        filterArr[o].Title.push(this.dataset.title);
        filterArr[o].outcome.push(this.dataset.outcome);
      }

      $.each(filterArr[o].arr, function (i, el) {
        if ($.inArray(el, filterArr[o].Finalarr) === -1)
          filterArr[o].Finalarr.push(el);
      });

      $.each(filterArr[o].Title, function (i, el) {
        if ($.inArray(el, filterArr[o].FinalTitle) === -1)
          filterArr[o].FinalTitle.push(el);
      });
    }
  });

  chr = 0;
  questionChars = [
    "A",
    "B",
    "C",
    "D",
    "E",
    "F",
    "G",
    "H",
    "I",
    "J",
    "K",
    "L",
    "M",
    "N",
    "O",
    "P",
    "Q",
    "R",
    "S",
    "T",
    "U",
    "V",
    "W",
    "X",
    "Y",
    "Z",
  ];

  for (var o in filterArr) {
    filterArr[o].ID.push(
      String(filterArr[o].type).substr(
        0,
        String(filterArr[o].type).indexOf(" ")
      )
    );

    $("#pdfPreview #sContainIn").append(
      '<h2 id = "main_Head" class =' + filterArr[o].ID + "_h2></h2>"
    );

    for (var index = 0; index < filterArr[o].Finalarr.length; index++) {
      $("#pdfPreview #sContainIn").append(
        '<div class = "main_Div" id =' +
          (filterArr[o].ID +
            "_" +
            filterArr[o].Finalarr[index].split(" ").join("_")) +
          "></div>"
      );

      for (var k = 0; k < filterArr[o].FinalTitle.length; k++) {
        //if(k == 0){chr = '1';}else{
        //}

        $(
          "#" +
            filterArr[o].ID +
            "_" +
            filterArr[o].Finalarr[index].split(" ").join("_")
        ).append(
          "<div style='display:flex;justify-content:space-between;'><h3 style='width:90%;'>" +
            questionChars[chr] +
            ". &nbsp; &nbsp;" +
            filterArr[o].FinalTitle[k] +
            "</h3><h3 class='outcomeText'>" +
            filterArr[o].outcome[k] +
            "</h3></div>"
        );

        chr = nextChar(chr);
        totalCat_marks = 0;

        $("#sContainIn .question-parent").each(function (i, e) {
          if ($(this).attr("data-type") == filterArr[o].type) {
            totalCat_marks += parseInt(this.dataset.marks);

            if (filterArr[o].FinalTitle[k] == this.dataset.title) {
              $(
                "#" +
                  (filterArr[o].ID +
                    "_" +
                    filterArr[o].Finalarr[index].split(" ").join("_"))
              ).append($(this));
            }
          }
        });
      }
      break;
    }

    if (totalCat_marks > 1) {
      $("." + filterArr[o].ID + "_h2").html(
        filterArr[o].type /*+ ' &mdash; ' + totalCat_marks + ' Marks'*/
      );
    } else {
      $("." + filterArr[o].ID + "_h2").html(
        filterArr[o].type /*+ ' &mdash; ' + totalCat_marks + ' Mark'*/
      );
    }
  }

  $(".marksInput").on("focus", handleMarksInputF);
  $(".marksInput").on("blur", handleMarksInputB);
  //$('.rx-editable').on('dblclick', enterEdit);

  if ($("#pdfPreview #unseencomprehension").length > 0) {
    $("#pdfPreview #unseencomprehension")
      .detach()
      .insertAfter("#pdfPreview #userInfo");
  }

  setNumbering();
  showCart();

  $("#userInfo").parent().children().show().end().end().hide();

  if (document.getElementById("student-version").checked == false) {
    $(".question-answer").hide();
  } else {
    $(".question-answer").show();
  }

  jQuery("#sContainIn .topTitlez").each(function (e) {
    if (!jQuery(this).find(".question-parent").length) jQuery(this).remove();
  });

  if (document.getElementById("student-version-m").checked == false) {
    $(".marksDiv").hide();
  }

  $("#auto-level-1").hide();
  $("#auto-level-2").show();
  setHeightFix();
  prepareSummary();
  disableReplaceBtn();
}

function showPreviewAuto() {
  loadImage();
  $("#auto-level-2").hide();
  $("#auto-level-3").show();

  $("#previewHTML")
    .html($("#pdfPreview").html())
    .promise()
    .done(function () {
      var school = $("#pdfPreview #school").val();
      $("#previewHTML #school").replaceWith(
        '<span id="text-school">' + school + "</span>"
      );

      var subject = $("#pdfPreview #subject").val();
      $("#previewHTML #subject").replaceWith(
        '<span id="text-subject">' + subject + "</span>"
      );

      var class_t = $("#pdfPreview #class_t").val();
      $("#previewHTML #class_t").replaceWith(
        '<span id="text-class_t">' + class_t + "</span>"
      );

      var name = $("#pdfPreview #name").val();
      $("#previewHTML #name").replaceWith(
        '<span id="text-name">' + name + "</span>"
      );

      // var roll = $('#pdfPreview #roll').val();
      // $('#previewHTML #roll').replaceWith('<span id="text-roll">' + roll + '</span>');

      var time = $("#pdfPreview #time").val();
      $("#previewHTML #time").replaceWith(
        '<span id="text-time">' + time + "</span>"
      );

      // var set = $('#pdfPreview #set').val();
      // $('#previewHTML #set').replaceWith('<span id="text-set">' + set + '</span>');

      $("#previewHTML #showPrevBtn").remove();

      var totalMarks = 0;
      $("#previewHTML .marksInput").each(function (i, e) {
        //get mark
        var spcMark = $("#pdfPreview .marksInput").eq(i).val();
        totalMarks += parseInt(spcMark);

        var $marks = $('<div class="marksPreview"></div>').html(
          "[Marks: " + spcMark + "]"
        );
        $(this).replaceWith($marks);
      });

      $("<div></div>")
        .append(
          '<label>Total Marks:</label><span id="text-marks">' +
            totalMarks +
            "</div>"
        )
        .appendTo("#previewHTML #userInfo");
      $("#previewHTML #userInfo .clear")
        .detach()
        .appendTo("#previewHTML #userInfo");

      var $table = $("<table></table>");
      $table.append("<tr><td></td><td></td><tr>");
      $table.append("<tr><td></td><td></td><tr>");
      $table.append("<tr><td></td><td></td><tr>");

      var subjectobj =
        '<tr><td align="left" valign="top"><label class="topHeadLable">Subject:</label></td><td class="topHeadTd" align="left" valign="top"><span id="text-subject">' +
        subject +
        "</span></td></tr>";
      var classobj =
        '<tr><td align="left" valign="top"><label class="topHeadLable">Class:</label></td><td class="topHeadTd" align="left" valign="top"><span id="text-class_t">' +
        class_t +
        "</span></td></tr>";
      var nameobj =
        '<tr><td align="left" valign="top" width="100px"><label class="topHeadLable">Name:</label></td><td class="topHeadTd" align="left" valign="top"><span id="text-name">' +
        name +
        "</span></td></tr>";
      //var rollobj = '<tr><td align="left" valign="top"><label class="topHeadLable" style="width:73px";>Roll No:</label></td><td  class="topHeadTd" align="left" valign="top"><span id="text-roll">' + roll + '</span></td></tr>';

      var schoolobj =
        '<tr><td align="left" valign="top"><label class="topHeadLable">School:</label></td><td class="topHeadTd" align="left" valign="top"><span id="text-school">' +
        school +
        "</span></td></tr>";
      //var setobj = '<tr><td align="left" valign="top"><label class="topHeadLable" style="width:69px";>Set no:</label></td><td class="topHeadTd" align="left" valign="top"><span id="text-set">' + set + '</span></td></tr>';
      var timeobj =
        '<tr><td align="left" valign="top"><label class="topHeadLable">Time:</label></td><td class="topHeadTd" align="left" valign="top"><span id="text-time">' +
        convertMinsToHrsMins(time) +
        "</span></td></tr>";
      var marksobj =
        '<tr><td align="left" width="150px" valign="top"><label class="topHeadLable" style="width:118px";>Total Marks:</label></td><td class="topHeadTd" align="left" valign="top"><span id="text-marks">' +
        totalMarks +
        "</span></td></tr>";

      if (testType == "worksheet") {
        marksobj = "";
      }

      var tablestr =
        '<table width="100%">\
		<tr>\
			<td width="50%" align="left" valign="top"><table>' +
        schoolobj +
        classobj +
        nameobj +
        '</table></td>\
			<td width="50%" align="left" valign="top"><table>' +
        subjectobj +
        timeobj +
        marksobj +
        "</table></td>\
		</tr>\
	</table>";

      $table.appendTo("#userInfo");
      $("#previewHTML #userInfo").html(tablestr);

      var stdnt = document.getElementById("student-version");
      if (stdnt.checked == false) {
        //$('#previewHTML .question-answer').remove();
      }

      //$('#previewHTML h3').remove();
      $("#previewHTML").children().show();

      var num;
      $("#previewHTML img").each(function (i, e) {
        //var width = getImgWidth(this.src);
        //this.setAttribute('width', 600 + 'px');
        //$(this).parent().append('<span class="numbering">' + num + '</span> ');
      });

      $("#previewHTML .mx-tools").remove();
      $(".marksDiv span").remove();

      //$('#previewHTML').next().next().children().eq(0).prepend('1.');
      $("#previewHTML .numbering").each(function (i, e) {
        //num = $(this).text();
        //$(this).next().next().children().eq(0).prepend('<span class="numbering">' + num + '</span> ');
        //$(this).remove();
      });

      //$('#previewHTML').prepend(userImageLogo);

      // $("#previewHTML").prepend('<div style="opacity: 0;">&nbsp;</div>');
      setNumberingPreview();
    });
  $(
    "#previewHTML .mx-tools, #previewHTML .question-info, .summary-selection"
  ).remove();

  $("#previewHTML").height("77%");
}

function autoNewQuestion() {
  $("#aft-validation").hide();
  $("#insertNewQuestion input[type='text']").val("");
  $("#insertNewQuestion").fadeIn();

  var questionType = allQuestions().distinct("type");
  var typeHTML = "";
  for (var i = 0; i < questionType.length; i++) {
    typeHTML +=
      '<option value="' +
      questionType[i] +
      '">' +
      questionType[i] +
      "</option>";
  }
  typeHTML += '<option value="' + "others" + '">' + "Others" + "</option>";
  $("#afm-type").html(typeHTML).trigger("change");

  var questionOutcome = allQuestions().distinct("outcome");
  var outcomeHTML = "";
  for (var i = 0; i < questionOutcome.length; i++) {
    outcomeHTML +=
      '<option value="' +
      questionOutcome[i] +
      '">' +
      questionOutcome[i] +
      "</option>";
  }
  $("#aft-outcome").html(outcomeHTML);
}

function switchScreenAuto(ele) {
  $(".auto-level-page")
    .not("#" + ele)
    .hide();
  $("#" + ele).show();
  setHeightFix();
}

function autoNewQuestionAdd() {
  $("#aft-validation").hide();

  var formObj = {
    type: document.getElementById("afm-type"),
    title: document.getElementById("afm-question-title"),
    question: document.getElementById("afm-question"),
    answer: document.getElementById("afm-answer"),
    marks: document.getElementById("afm-marks"),
    difficulty: document.getElementById("aft-difficulty"),
    outcome: document.getElementById("aft-outcome"),
  };

  var titleHelper = document.getElementById("afm-type-custom");

  var showMsg = function (msg) {
    $("#aft-validation").text(msg).show();
  };

  /*
	if(formObj.type.value == "others" && titleHelper.value == "") {
		showMsg("Please type type");
		formObj.type.focus();
		return false;
	}

	if(formObj.type.value == "others") {
		formObj.type = titleHelper;
	}
	*/

  if (formObj.title.value == "") {
    showMsg("Please type title");
    formObj.title.focus();
    return false;
  }

  if (formObj.question.value == "") {
    showMsg("Please type question");
    formObj.question.focus();
    return false;
  }

  if (formObj.answer.value == "") {
    showMsg("Please type answer");
    formObj.answer.focus();
    return false;
  }

  if (formObj.marks.value == "") {
    showMsg("Please insert marks");
    formObj.marks.focus();
    return false;
  }

  if (isNaN(formObj.marks.value)) {
    showMsg("Please insert valid marks in integer");
    formObj.marks.focus();
    return false;
  }

  var newId = String(allQuestions().count() + 1);
  var questionData = formObj.question.value.split("\n");
  var questionTitleMain = formObj.title.value;
  var questionOpts = [];

  var formattedObj = {
    class: "Custom Other",
    chapter: "Custom Other",
    neworbook: "Custom Other",
    topic: "Custom Other",
    // type: "others",
    type: formObj.type.value,
    question: {
      title: questionTitleMain,
      option: questionData,
      image: "",
    },
    answer: [formObj.answer.value],
    mark: formObj.marks.value,
    outcome: formObj.outcome.value,
    difficulty: formObj.difficulty.value,
    id: parseInt(newId),
  };

  if (document.getElementById("afm-image").files.length > 0) {
    formattedObj["question"]["image"] = inputToURL(
      document.getElementById("afm-image")
    );
  }

  allQuestions.insert(formattedObj);

  testBank.push(newId);

  if (tgType == "manual") {
    addToTest();
  } else {
    addToTestAuto(testBank);
  }

  dismissPopupInsertion();
  showMsgBox("New question has been added.");
  return false;
}

function dismissPopupInsertion() {
  $("#insertNewQuestion").hide();
  $("#insertNewQuestion input[type='text'], #insertNewQuestion textarea").val(
    ""
  );
  $("#new-question-insertion-form")[0].reset();
}

function inputToURL(inputElement) {
  var file = inputElement.files[0];
  return window.URL.createObjectURL(file);
}

var style =
  "<style>\
.question-parent {\
clear: both;\
margin: 0px;\
padding: 6px 0px 6px 0px;\
position: relative;\
margin-top: 7px;min-height: 22px;}\
.question-parent.selection{\
}\
.numbering {float: left;}\
.marksDiv{float: right; widht: 25px;}\
</style>";

function printStudent() {
  $("#previewHTML .question-answer").hide();
  if (typeof appjs != "object") {
    if (isFirefox) {
      //exportPDFKendo('student');
      printCommand();
    } else {
      printCommand();
    }
  } else {
    if (preservemyHTML == "") {
      preservemyHTML = $("#previewHTML").html();
    }
    $("#previewHTML").html(preservemyHTML);
    $("#previewHTML .question-answer").remove();
    devilzgetpdf(style + $("#previewHTML").html());
  }
}

function printTeacher() {
  $("#previewHTML .question-answer").show();
  if (typeof appjs != "object") {
    if (isFirefox) {
      //exportPDFKendo('teacher');
      printCommand();
    } else {
      printCommand();
    }
  } else {
    if (preservemyHTML == "") {
      preservemyHTML = $("#previewHTML").html();
    }
    $("#previewHTML").html(preservemyHTML);

    devilzgetpdf(style + $("#previewHTML").html());
  }
  $(window).focus(function () {
    console.log("focused");
  });
}

function exportPDFKendo(type) {
  var printHTML = $("#previewHTML");
  var height = printHTML.height();
  printHTML.height("auto");
  var filename = type + "_" + moment().format("DD-MM-YYYY-hhmmss") + ".pdf";

  var printType = "all";

  if (printType == "kendo1") {
    kendo.drawing.drawDOM(printHTML).then(function (group) {
      kendo.drawing.pdf.saveAs(group, filename);
      printHTML.height(height);
    });
  } else {
    $("body").addClass("printsection");
    var pdfOptions = {
      forcePageBreak: ".page-break",
      margin: { left: "1cm", top: "1cm", right: "1cm", bottom: "1cm" },
      paperSize: "a4",
      multiPage: true,
    };

    kendo.drawing
      .drawDOM(printHTML, pdfOptions)
      .then(function (group) {
        return kendo.drawing.exportPDF(group, pdfOptions);
      })
      .done(function (data) {
        printHTML.height(height);
        $("body").removeClass("printsection");
        kendo.saveAs({
          dataURI: data,
          fileName: filename,
          proxyURL: "https://demos.telerik.com/kendo-ui/service/export",
        });
      });
  }
}

function generatePDF() {
  var specialElementHandlers = {
    "#bypassme": function (element, renderer) {
      return true;
    },
  };
  var tinymceToJSPDFHTML = $("#previewHTML").html();
  var margin = {
    top: 30,
    left: 45,
    right: 30,
    bottom: 45,
  };
  var doc = new jsPDF("p", "pt", "letter");
  var canvas = doc.canvas;

  canvas.width = 8.5 * 72;

  doc.fromHTML(
    tinymceToJSPDFHTML,
    0,
    0,
    {
      width: 568, // max width of content on PDF
      elementHandlers: specialElementHandlers,
    },
    function (bla) {
      doc.save("saveInCallback.pdf");
    },
    margin
  );
}

window.onafterprint = function () {
  $("#css-responsive").attr("href", "style/responsive1.css");
};

window.onbeforeprint = function () {
  $("#css-responsive").attr("href", "");
};
