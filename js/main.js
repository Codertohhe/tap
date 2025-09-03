var $treeview;
var testBank = new Array();
var queryTrack = new Object();
var parentsids;
var parentsobjHandle = {};
var selectChapters = [];
const EDIF = {
  E: "Easy",
  M: "Medium",
  H: "Hard",
};

var filterArr = [];
var chr = 0;

$(document).ready(function (e) {
  $(".tgTitle").text(tgTitle);
  $('a[href="#"]').click(function (e) {
    e.preventDefault();
  });

  $("#afm-type").on("change", function (e) {
    if (this.value == "others") {
      $("#afm-type-input").show();
    } else {
      $("#afm-type-input").hide();
    }
  });

  /*$.mCustomScrollbar.defaults.theme="light-2"; //set "light-2" as the default theme
	$("#container").mCustomScrollbar({
		axis:"yx"
	});*/

  $treeview = $("#collectionList #container");

  $treeview.append($("#treeTemplate").html());

  $treeview.jstree({
    core: {
      multiple: false,
    },
  });

  $treeview.jstree("open_all");

  $treeview.on("open_node.jstree", function (e, data) {
    selectChapters = uniqueArray(selectChapters);
    $.each(selectChapters, function (i, e) {
      //console.log(selectChapters[i]);
      $("#" + selectChapters[i]).addClass("practive");
    });

    // adding title for tree a
    $("#collectionList a").each(function (e) {
      var text = this.textContent;
      this.setAttribute("title", text);
    });
  });

  $treeview.on("changed.jstree", function (e, data) {
    setScrollbar();

    //get selected element id
    var selectedId = data.selected[0];
    //console.log('selected: ' + selectedId);
    handleStructure(selectedId);
  });

  setActionForSavingMarks();
});

function getUniqueData(keyname, keyvalue) {
  return allQuestions().distinct(keyname);
}

function handleStructure(selectedId) {
  //select node
  $(".jstree-clicked").removeClass("jstree-clicked");
  $("#" + selectedId + " > a").addClass("jstree-clicked");

  //open the tree on being selected
  $treeview.jstree("open_node", selectedId);

  //get paths for breakcrums
  var parents = new Array();
  var parentsLink = new Array();
  parents.push(
    $("#" + selectedId)
      .children("a")
      .text()
  );
  $("#" + selectedId)
    .parents("li")
    .each(function (i, e) {
      parentsLink.push(this.id);
      var txt = $(this).children("a").text();
      parents.push(txt);
    });

  //collect ids for highlight selection
  parentsids = [];
  $("#" + selectedId)
    .parents("li")
    .each(function (i, e) {
      parentsids.push(this.id);
    });
  parentsids.reverse();
  parentsLink.reverse();
  if (parentsids[1] != undefined && parents[parents.length - 2] != undefined)
    parentsobjHandle[parentsids[1]] = parents[parents.length - 2];
  $("#explorer").empty();

  $('<div id="breadcrums"></div>').appendTo("#explorer");
  $("#breadcrums").append(
    "<span class='breadLabel'>" +
      $(parents)
        .get()
        .reverse()
        .join(
          "</span> <span class='fa fa-chevron-right'></span> <span class='breadLabel'>"
        ) +
      "</span>"
  );
  // enable breadcrums link
  $("#breadcrums .breadLabel").each(function (i, e) {
    $(this).attr("data-link", parentsLink[i]);
  });
  //check if it has further children
  if ($("#" + selectedId + " > ul").length >= 1) {
    $("#explorer").append('<div id="folders"></div>');
    $("#" + selectedId + " > ul > li > a").each(function (i, e) {
      function showBox(e) {
        e.preventDefault();
        //alert($(this).text());
        var itemId = this.dataset.id;
        handleStructure(itemId);
      }

      var text = $(this).text();
      var $folder = createFolder(text);

      $folder
        .click(showBox)
        .attr("data-id", $(this).parent().attr("id"))
        .appendTo("#folders");
    });
    setHeightFix();
  } else {
    console.log("question being inserted");
    var scrollPos = $("#level1_explorer").offset().top;
    $("body").animate({ scrollTop: scrollPos });
    var treeLength = treeStructure.length;
    var myLink = new Array();
    var query_question = new Object();

    for (var i = 0; i < treeLength; i++) {
      myLink[i] = parents[i].trim();
      query_question[treeStructure[treeLength - i - 1]] = myLink[i];
    }

    var formattedQuery = JSON.stringify(query_question);
    var totalQuestion = allQuestions(query_question);
    var filterDifficulty = $("#manual-difficulty").val();
    var filterOutcome = $("#manual-outcome").val();
    if (filterOutcome == null) filterOutcome = [];
    if (filterDifficulty == null) filterDifficulty = [];

    $("#explorer").append(
      '<div id="stats"><div><div for="totalQuestions"><span class="quesSec">Questions selected in this section: </span><span class="quesSecNo"><input type="text" name="questionSelected" id="questionSelected" readonly="" value="0">/<input type="text" name="totalQuestions" id="totalQuestions" readonly="" value="0"></span></div></div><button class="info_btn" id="section-prev-test" onclick="addAndMoveOn();">Preview</button><button class="info_btn" id="addTest" onclick="addToTest();">Add to test</button><button class="info_btn" id="selectAll" onclick="selectAllQuestion();">Select All</button>'
    );

    $("#explorer").append('<div id="questionsViewer"></div>');

    $("#totalQuestions").val(totalQuestion.count());
    var totalPreSelected = 0;
    var counter = 0;
    var activeChapters = [];
    totalQuestion.each(function (r) {
      var indexDifficulty = filterDifficulty.indexOf(r.difficulty);
      var indexOutcome = filterOutcome.indexOf(r.outcome);

      //if(r.type.toLowerCase() == "Fill in the blanks." || r.type.toLowerCase() == "short answer questions" || r.type.toLowerCase() == "long answer questions")
      if (indexDifficulty != -1 && indexOutcome != -1) {
        counter++;

        var $div = $('<div class="question-parent"></div>').attr({
          id: "fib",
          "data-type": r.type,
          "data-cid": r.id,
          "data-ref": formattedQuery,
        });
        var $checkbox = $("<input/>").attr({
          type: "checkbox",
          id: "check" + counter,
        });
        queryTrack[String(r.id)] = formattedQuery;
        var a = $.inArray(String(r.id), testBank);
        if (a == -1) {
          $checkbox.appendTo($div).click(function (e) {
            $div.toggleClass("selection");
            $("#questionSelected").val($(".selection").length);
          });
        } else {
          totalPreSelected++;
          $div.addClass("selection dis");
          $checkbox
            .appendTo($div)
            .click(function (e) {})
            .prop("checked", true)
            .attr("disabled", "disabled")
            .css("cursor", "default");
        }

        var bulleting = "";
        if (r.type.toLowerCase() == "mcqaaaaaaa") {
          bulleting = "● ";
        }
        var $opts;
        for (var q in r.question) {
          if (q == "title") {
            var flag = $("#questionsViewer .hTitle").filter(function (i, e) {
              //console.log(e.innerHTML, r.question.title, this.innerHTML == r.question.title)
              return this.textContent == r.question.title;
            });

            //console.log($('#questionsViewer .hTitle').text(), $('#questionsViewer .hTitle').html());
            //console.log(flag, ' flag.length')
            if (flag.length == 0) {
              $('<div class="question-title hTitle"></div>')
                .html(/*counter + '. ' + */ r.question.title)
                .appendTo($div);
            }
          } else if (q == "image") {
            if (
              r.question.image != undefined &&
              r.question.image != "" &&
              r.question.image != "no_img"
            )
              $('<div class="image"></div>')
                .html('<img src="' + imagePath + r.question.image + '" />')
                .appendTo($div);
          } else if (q == "option") {
            $opts = $('<div class="options"></div>');

            if (r.type.toLowerCase() == "mtfaaaaaaaa") {
              var optnameii =
                "<br><table border='1' cellspacing='0' cellpadding='0'>";
              for (
                var mtfIndex = 0;
                mtfIndex < r.question.option.length;
                mtfIndex++
              ) {
                var optnameiis = r.question.option[mtfIndex];
                var mtfoption = optnameiis.split(" ^ ");

                optnameii +=
                  "<tr><td>" +
                  mtfoption[0] +
                  "</td><td>" +
                  mtfoption[1] +
                  "</td></tr>";
              }
              optnameii += "</table>";

              $('<div class="option rx-editable"></div>')
                .html(optnameii)
                .appendTo($opts);
            } else {
              /*if(r.type.toLowerCase() == "fill in the blanks")
							{
								for(var z=0; z<r.question.option.length; z++)
								{
									var optnameii = r.question.option[z];
									
									console.log(optnameii.indexOf('|'), " optnameii")
									if(optnameii.indexOf("|") != -1)
										{
											var head = optnameii.substr(0,optnameii.indexOf("|"));
											var head_2 = optnameii.substr((parseInt(optnameii.indexOf("|"))+1),optnameii.length);
											$('<div class="question-title"></div>').html('<span class="head">'+head+'</span>' + '<span class="head_2">'+head_2+'</span>').appendTo($opts);										
										//optnameii = "<img src='" + imagePath + r.question.option[z] + "'>"
										}
									else $('<div class="question-title"></div>').html(bulleting + optnameii).appendTo($opts);
								}
							}
							else
							{
							*/
              for (var z = 0; z < r.question.option.length; z++) {
                var optnameii = r.question.option[z];

                if (optnameii.indexOf("<img src") == -1) {
                  if (
                    optnameii.indexOf(".jpg") != -1 ||
                    optnameii.indexOf(".png") != -1
                  ) {
                    optnameii =
                      "<img src='" + imagePath + r.question.option[z] + "'>";
                  }
                }

                $('<div class="question-title"></div>')
                  .html(bulleting + optnameii)
                  .appendTo($opts);
                //$('<div class="question-title"></div>').html(String.fromCharCode(65+z) + '. ' + r.question.option[z]).appendTo($div);
              }
              //}
            }
          }
        }
        $opts.appendTo($div);

        $("#questionSelected").val(totalPreSelected);
        $('<div class="clear"></div>').appendTo($div);
        $("#questionsViewer").append($div);

        $div.append(
          "<div class='question-outcome'>Source: " + r.outcome + "</div>"
        );
        $div.append(
          "<div class='question-difficulty'>Difficulty: " +
            EDIF[r.difficulty] +
            "</div>"
        );
        setHeightFix();
      }
    });
    $("#totalQuestions").val(counter);
    //console.log(counter);

    /** removing multiple htitle **/
    var tmph = "";
    $("#questionsViewer")
      .find(".hTitle")
      .each(function (i, e) {
        var ht = $(this).html();
        if (ht == tmph) {
          $(this).remove();
        }
        tmph = ht;
      });
  }
}

//create folder structure
function createFolder(folderName) {
  var $folder = $("<div></div>").attr("class", "main");
  $("<div></div>")
    .attr("class", "folder")
    .append('<div class="front"></div><div class="back"></div>')
    .appendTo($folder);
  $("<div></div>")
    .attr("class", "folder_name")
    .append('<a href="#">' + folderName + "</a>")
    .appendTo($folder);
  return $folder;
}

var totalMarksSummary = 0;
var difficultySummary = {
  E: 0,
  M: 0,
  H: 0,
};
var outcomeSummary = {};

function addToTest() {
  var selectedQuestions = $("#questionsViewer .selection")
    .map(function (index, element) {
      return this.dataset.cid;
    })
    .get();

  //console.log(selectedQuestions);
  //$('#questionsViewer :checked').prop('checked', false);

  for (var i = 0; i < selectedQuestions.length; i++)
    testBank.push(selectedQuestions[i]);

  //testBank.sort(function(a, b){return (a-b);});
  testBank = uniqueArray(testBank);

  /*
	var bankDB = []
	allQuestions().order("type").each(function(record, recordNumber) {
		if(testBank.indexOf(record.id.toString()) != -1)
			bankDB.push(record.id.toString());
	});
	testBank = bankDB;
	*/

  //console.log('bankdb', bankDB);
  //$('#cartItems').text('(' + testBank.length + ')');
  $("#cartItems").text(testBank.length);

  $("#pdfPreview").empty();

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
  // $userInfo.append('<div id="form-roll"><label>Roll no:</label> <input type="text" name="roll" id="roll" style="margin-left: 1px;margin-top: 0px;}" /></div>');
  $userInfo.append(
    '<div id="form-name"><label>Name:</label> <input type="text" name="name" id="name" style="margin-left: 14px;" /></div>'
  );
  $userInfo.append(
    '<div id="form-time"><label>Time:</label> <input type="text" name="time" onkeydown="return numValidate(event)" id="time" style="margin-left: 12px; margin-top: -3px;" /></div>'
  );
  // $userInfo.append('<div id="form-set"><label>Set no:</label> <input type="text" name="set" id="set" style="margin-left: 7px;" /></div>');
  $userInfo.append(
    '<div id="logoUpdateParent" style="width: 345px;"><label>Logo </label> <input type="file" name="imgfile" id="imgfile" accept="image/*" size="50"/><div class="fileChooseBlocker"></div></div>'
  );
  $userInfo.append(
    '<div id="form-btn"><button id="showPrevBtn" class="sm_button" onclick="showPreview()">Submit</button></div>'
  );
  $userInfo.append(
    '<button id="closepopupBtn" class="" title="Close" onclick="closePopupFromPreview()"><i class="fa fa-times"></i></i></button>'
  );

  //console.log("details", $userInfo);

  $userInfo.append('<div class="clear"></div>');

  $("#pdfPreview").append($userInfo);
  testBank = uniqueArray(testBank);
  //console.log(testBank);

  selectChapters = [];
  $(".practive").removeClass("practive");

  difficultySummary = {
    E: 0,
    M: 0,
    H: 0,
  };
  outcomeSummary = {};
  totalMarksSummary = 0;
  for (var i = 0; i < testBank.length; i++)
    insertQuestion($("#pdfPreview"), testBank[i]);

  for (var zz = 0; zz < selectChapters.length; zz++) {
    $("#" + selectChapters[zz]).addClass("practive");
  }

  var categoriges = $("#pdfPreview #sContainIn")
    .children()
    .map(function () {
      return this.dataset.title;
    })
    .get();
  var chapters = $("#pdfPreview #sContainIn")
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

  chapters = uniqueArray(chapters);
  categoriges = uniqueArray(categoriges);
  //console.log(chapters);
  //console.log(categoriges);

  var uniquetype = [];
  $.each(type, function (i, el) {
    if ($.inArray(el, uniquetype) === -1) uniquetype.push(el);
  });

  /* 	For filtering and merging questions by type */

  var $section;
  var totalCat_marks;
  /*
   * For filtering and merging questions by type
   */
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
  //showCart();

  $("#userInfo").parent().children().show().end().end().hide();

  if (document.getElementById("student-version").checked == false) {
    $(".question-answer").hide();
  } else {
    $(".question-answer").show();
  }

  jQuery("#sContainIn .topTitlez").each(function (e) {
    if (!jQuery(this).find(".question-parent").length) jQuery(this).remove();
  });
  if (testBank.length == 0) {
    switchScreen("level1");
  }

  if (document.getElementById("student-version-m").checked == false) {
    $(".marksDiv").hide();
  }

  prepareSummary();
  disableReplaceBtn();
}

function actionReplaceBtn() {
  $(".tool-replace").each(function (i, e) {
    var replacerId = $(this).parent().parent().attr("id");
    //var oldQuery = JSON.parse(queryTrack[replacerId]);
    $("#replaceList").empty();
    var $quesContainer = $(this).parent().parent();
    var subject = $quesContainer.attr("data-subject");
    var chapter = $quesContainer.attr("data-chapter");

    var marks = $quesContainer.attr("data-marks");

    var topic = $quesContainer.attr("data-topic");
    var type = $quesContainer.attr("data-type");

    var oldQuery = new Object();

    if (type != undefined) oldQuery["type"] = type;
    if (topic != undefined) oldQuery["topic"] = topic;
    if (chapter != undefined) oldQuery["chapter"] = chapter;
    if (subject != undefined) oldQuery["subject"] = subject;

    if (marks != undefined) oldQuery["mark"] = marks;

    var totalQuestion = allQuestions(oldQuery);

    uniqueQues = [];
    totalQuestion.each(function (r) {
      var idq = String(r.id);
      if (testBank.indexOf(idq) == -1) {
        uniqueQues.push(idq);
      }
    });
    if (uniqueQues.length == 0) {
      $(this).attr("disabled", "disabled");
    }
  });
}

function disableReplaceBtn() {
  return false;
  $("#level2 #pdfPreview").hide();
  $("#auto-level-2 #pdfPreview").hide();
  setTimeout(function () {
    $("#level2 #pdfPreview").show();
    $("#auto-level-2 #pdfPreview").show();
    actionReplaceBtn();
  }, testBank.length * 500);
}

function prepareSummary() {
  var summary = {};
  summary["questions"] = testBank.length;
  summary["marks"] = totalMarksSummary;
  summary["difficulty"] = [];
  summary["outcome"] = [];

  if (difficultySummary["E"] != 0) {
    summary["difficulty"].push("Easy: " + difficultySummary["E"]);
  }
  if (difficultySummary["M"] != 0) {
    summary["difficulty"].push("Medium: " + difficultySummary["M"]);
  }
  if (difficultySummary["H"] != 0) {
    summary["difficulty"].push("Hard: " + difficultySummary["H"]);
  }

  for (var key in outcomeSummary) {
    if (
      outcomeSummary[key] == undefined ||
      outcomeSummary[key] == null ||
      outcomeSummary[key] == ""
    ) {
      delete outcomeSummary[key];
    } else {
      summary["outcome"].push(key + ": " + outcomeSummary[key]);
    }
  }

  $("#pdfPreview").append(
    "" +
      '<div class="summary-selection">' +
      '<div class="summary-container">' +
      '<div class="summary_leftpane">' +
      '<div class="summary-total">' +
      "<p>Questions: <span>" +
      summary["questions"] +
      "</span></p>" +
      "</div>" +
      '<div class="summary-marks">' +
      "<p>Marks: <span>" +
      summary["marks"] +
      "</span></p>" +
      "</div>" +
      '<div class="summary-diffictulty">' +
      '<p>Difficulty: <span class="outcome-wrap">' +
      summary["difficulty"].join("</span>, <span class='outcome-wrap'>") +
      "</span></p>" +
      "</div>" +
      '<div class="summary-outcome">' +
      '<p>Source: <span class="outcome-wrap">' +
      summary["outcome"].join("</span>, <span class='outcome-wrap'>") +
      "</span></p>" +
      "</div>" +
      "</div>" +
      '<div class="summary_rightPane">' +
      '<div class="macm_logo">' +
      //'<img src="img/logo-macmillan.png">' +
      "</div>" +
      "</div>" +
      "</div>" +
      "</div>"
  );
}

function toggleAnswers() {
  if (document.getElementById("student-version").checked == false) {
    $(".question-answer").hide();
  } else {
    $(".question-answer").show();
  }
}
function toggleMarks() {
  if (document.getElementById("student-version-m").checked == false) {
    $(".marksDiv").hide();
  } else {
    $(".marksDiv").show();
  }
}
function print_option() {
  $("#print_option").show();
}
function showCart() {
  $("#level1").hide();
  $("#level2").show();
  $("#showCartBtn").hide();
  setHeightFix();
}

function enterEdit(e) {
  $(this).attr({ contenteditable: true, spellcheck: false }).focus();
  //.off('dblclick');
  $(this).on("keydown", handleEnter);
}

function handleEnter(e) {
  if (e.keyCode == 13) {
    $(this).blur().attr("contenteditable", false);
    //.on('dblclick', enterEdit);
    return false;
  }
}

function handleMarksInputF(e) {
  if (this.value == "0") this.value = "";
}

function handleMarksInputB(e) {
  if (this.value == "") this.value = "0";

  var givenMarks = this.value;
  var quesID = $(this).parents(".question-parent").attr("id");
  allQuestions({ id: parseInt(quesID) }).each(function (r) {
    r.mark = parseInt(givenMarks);
  });

  //console.log(allQuestions({id:parseInt(quesID)}).get().mark);

  if (tgType == "auto") {
    addToTestAuto(testBank);
  } else {
    addToTest();
  }
}

function escapeRegExp(string) {
  return string.replace(/([.*+?^=!:${}()|\[\]\/\\])/g, "\\$1");
}

function replaceAll(string, find, replace) {
  return string.replace(new RegExp(escapeRegExp(find), "g"), replace);
}

function nextChar(c) {
  //return String.fromCharCode(c.charCodeAt(0) + 1);
  //console.log(chr, c, ' before');
  chr++;
  //c = c++;
  //	console.log(chr, c, ' after');
  return chr;
}

function setNumbering() {
  var ind = 0;
  var numbering = 1;
  var chr = "a";
  var numnew = 1;
  $("#pdfPreview").each(function (i, e) {
    var prev_chp = "";
    var prev_typ = "";
    var prev_title = "";

    $(this)
      .find(".sContainIn")
      .each(function (i, e) {
        var tit = "<h3>" + $(this).find("h3").html() + "</h3>";

        var len = $(this).find(".question-parent").length;
        if (len > 1) {
          var cc = 1;
          $(this)
            .find(".question-parent")
            .each(function (i, e) {
              var chp = $(this).attr("data-chapter");
              var typ = $(this).attr("data-type");
              var title = $(this).attr("data-title");
              //console.log('chp:'+chp+ ' prev_chp:'+prev_chp + 'typ:'+typ+ ' prev_typ:'+prev_typ + 'num:'+numbering);
              //console.log(prev_title, title, prev_typ, typ, ' typ');

              if (
                /*(prev_chp != chp) || */ prev_typ != typ ||
                prev_title != title
              ) {
                numbering = 1;
                //if(cc != 1){$(this).before(tit);}
              }
              prev_chp = chp;
              prev_typ = typ;
              prev_title = title;

              var id = parseInt($(this).attr("id"));

              //console.log(allQuestions({id: id}).first().question, ' gggggggg');

              if (
                allQuestions({ id: id }).first().question.image != "" ||
                allQuestions({ id: id }).first().question.image == "no_img"
              )
                $(this).find("label.numbering").html(" ");
              else
                $(this)
                  .find("label.numbering")
                  .html(numbering + ".");

              //if(numbering == 1){chr = 'a';}else{chr = nextChar(chr);}
              //$(this).find('label.numbering').html(chr + '.');
              numbering++;
              cc++;
            });
        }

        $(this)
          .find("h3")
          .each(function (i, e) {
            /*var tit1 = $(this).html();
				var myarr = tit1.split(".");
				newStr= tit1.replace(myarr[0], numnew);
				$(this).html(newStr);
				numnew++;*/
          });
      });
  });
}

function insertQuestion($container, ids) {
  allQuestions({ id: parseInt(ids) }).each(function (r) {
    var mark = parseInt(r.mark);
    if (typeof mark == "number") {
      totalMarksSummary += mark;
    }

    difficultySummary[r.difficulty]++;

    if (
      outcomeSummary[r.outcome] == undefined ||
      outcomeSummary[r.outcome] == null
    ) {
      outcomeSummary[r.outcome] = 1;
    } else {
      outcomeSummary[r.outcome]++;
    }

    var qID = replaceAll(r.type, " ", "").toLowerCase();
    var qID = "sContainIn";

    if (!document.getElementById("sContainIn"))
      $container.append('<div id="sContainIn" class = "sContainIn"></div>');

    if ($("#" + qID).length == 0) {
      $('<div class="topTitlez"></div>')
        .attr({ id: qID })
        .html("<h3>" + r.question.title + "</h3>")
        .appendTo("#sContainIn");
    }

    for (key in parentsobjHandle) {
      if (parentsobjHandle[key] == r.chapter) selectChapters.push(key);
    }

    var type = r.type;

    var $div = $(
      // '<div class="question-parent myqstype-' + r.type + '"></div>'
      '<div class="question-parent myqstype-"></div>'
    ).attr({
      id: r.id,
      "data-type": r.type,
      "data-chapter": r.chapter,
      "data-topic": r.topic,
      "data-subject": r.subject,
      "data-title": r.question.title,
      "data-marks": r.mark,
      "data-outcome": r.outcome,
    });

    $("<label></label>").addClass("numbering").appendTo($div);

    if (r.mark == null) r.mark = "0";

    if (r.mark > -1) {
      //console.log(' marks : '+r.mark);
      $inpMarks = $('<input class="marksInput" />').attr({
        type: "text",
        value: r.mark,
        maxlength: 2,
        onkeydown: "return numValidate(event)",
      });
      $('<div id="marksDiv-id" class="marksDiv"></div>')
        .append("<span>Marks</span>")
        .append($inpMarks)
        .appendTo($div);
    }

    var $opts;

    var bulleting = "";
    if (type.toLowerCase() == "mcqzzzzz" && r.topic != "Custom Otherzzzzz") {
      bulleting = '<span class="aligntop">●</span> ';
    }

    for (var q in r.question) {
      if (q == "title") {
        //$('<div class="question-title rx-editable"></div>').html(r.question.title).appendTo($div);
        $('<div class="question-title rx-editable"></div>')
          .html(r.question.option[0] != null ? r.question.option[0] : "")
          .appendTo($div);
      } else if (q == "image") {
        if (
          r.question.image != undefined &&
          r.question.image != "" &&
          r.question.image != "no_img"
        ) {
          var qimagepathblob = imagePath + r.question.image;
          if (r.question.image.indexOf("blob:") == 0) {
            qimagepathblob = r.question.image;
          }
          $('<div class="image"></div>')
            .html('<img src="' + qimagepathblob + '"/>')
            .appendTo($div);
        }
      } else if (q == "option") {
        $opts = $('<div class="options"></div>');
        //console.log(type, ' r.question.option.length');
        if (type == "MTFzzzzzz") {
          var optnameii =
            "<br><table border='1' cellspacing='0' cellpadding='0'>";
          for (
            var mtfIndex = 0;
            mtfIndex < r.question.option.length;
            mtfIndex++
          ) {
            var optnameiis = r.question.option[mtfIndex];
            var mtfoption = optnameiis.split(" ^ ");

            optnameii +=
              "<tr><td>" +
              mtfoption[0] +
              "</td><td>" +
              mtfoption[1] +
              "</td></tr>";
          }
          optnameii += "</table>";

          $('<div class="option rx-editable"></div>')
            .html(optnameii)
            .appendTo($opts);
        } else {
          /*else if(type == "Fill in the blanks")
				{
					for(var z=0; z<r.question.option.length; z++)
					{
						var optnameii = r.question.option[z];
						
						if(optnameii.indexOf("|") != -1)
							{
								var head = optnameii.substr(0,optnameii.indexOf("|"));
								var head_2 = optnameii.substr((parseInt(optnameii.indexOf("|"))+1),optnameii.length);
								//$('<div class="question-title"></div>').html('<span class="head">'+head+'</span>' + '<span class="head_2">'+head_2+'</span>').appendTo($opts);										
							//optnameii = "<img src='" + imagePath + r.question.option[z] + "'>"
							}
							//else $('<div class="question-title"></div>').html(bulleting + optnameii).appendTo($opts);
						}
				}*/
          //console.log(type, ' come not');
          if (r.question.option.length == 2) {
            var z = 0;
            var optnameii = r.question.option[z];
            if (
              optnameii.indexOf(".jpg") != -1 ||
              optnameii.indexOf(".png") != -1
            ) {
              optnameii =
                "<img src='" + imagePath + r.question.option[z] + "'>";
            }
            $('<div class="option rx-editable"></div>')
              .html(bulleting + optnameii)
              .appendTo($opts);
          } else if (r.question.option.length > 2) {
            for (var z = 1; z < r.question.option.length; z++) {
              //$('<div class="option rx-editable"></div>').html(String.fromCharCode(65+z) + '. ' + r.question.option[z]).appendTo($opts);

              var optnameii = r.question.option[z];
              if (
                optnameii.indexOf(".jpg") != -1 ||
                optnameii.indexOf(".png") != -1
              ) {
                optnameii =
                  "<img src='" + imagePath + r.question.option[z] + "'>";
              }
              $('<div class="option rx-editable"></div>')
                .html(bulleting + optnameii)
                .appendTo($opts);
            }
          }
        }
      }
    }
    $opts.appendTo($div);

    if (r.answer != "" && r.answer != null) {
      if (Object.prototype.toString.call(r.answer) === "[object Array]") {
        // $ans = $('<div class="question-answer rx-editable"></div>').html('<div class="ansHeading">Gllej:</div>');
        $ans = $('<div class="question-answer rx-editable"></div>').html(
          '<div class="ansHeading">Answer:</div>'
        );

        if (r.answer.length == 1) {
          var z = 0;
          var optnameii = r.answer[z];
          if (
            optnameii.indexOf(".jpg") != -1 ||
            optnameii.indexOf(".png") != -1
          ) {
            optnameii =
              "<img class='ansImg' src='" + imagePath + r.answer[z] + "'>";
          }
          optnameii = optnameii.split("^").join("-");
          $ans.append("<span class='simp-ans-text'>" + optnameii + "</span>");
        } else {
          for (var z = 0; z < r.answer.length; z++) {
            var optnameii = r.answer[z];
            if (
              optnameii.indexOf(".jpg") != -1 ||
              optnameii.indexOf(".png") != -1
            ) {
              optnameii =
                "<img class='ansImg' src='" + imagePath + r.answer[z] + "'>";
            }
            optnameii = optnameii.split("^").join("-");
            $ans.append("<span class='simp-ans-text'>" + optnameii + "</span>");

            //$ans.append('<div>' + String.fromCharCode(65+z) + '. ' + r.answer[z] + '</div>');
          }
        }
        $ans.appendTo($div);
      } else {
        $('<div class="question-answer rx-editable"></div>')
          //   .html('<div class="ansHeading">Gllej:</div> ' + r.answer)
          .html('<div class="ansHeading">Answer:</div> ' + r.answer)
          .appendTo($div);
      }
    }

    $div.append(
      "" +
        '<div class="question-info">' +
        '<div class="question-difficulty">Difficulty: <span>' +
        EDIF[r.difficulty] +
        "</span></div>" +
        '<div class="question-outcome">Source: <span>' +
        r.outcome +
        "</span></div>" +
        "</div>"
    );

    var $tools = $('<div class="mx-tools"></div>').appendTo($div);

    $('<a href="javascript:void(0)"></a>')
      .html("Replace")
      .attr({ class: "tool-replace" })
      .click(function (i, e) {
        if ($(this).attr("disabled") == "disabled") {
          return false;
        }
        var replacerId = $(this).parent().parent().attr("id");
        //var oldQuery = JSON.parse(queryTrack[replacerId]);
        $("#replaceList").empty();
        var $quesContainer = $(this).parent().parent();
        var subject = $quesContainer.attr("data-subject");
        var chapter = $quesContainer.attr("data-chapter");
        var topic = $quesContainer.attr("data-topic");
        var type = $quesContainer.attr("data-type");

        //console.log('subject: ' + subject + ' chapter: ' + chapter + ' topic: ' + topic + ' type: ' + type);
        var oldQuery = new Object();
        if (type != undefined) oldQuery["type"] = type;
        if (topic != undefined) oldQuery["topic"] = topic;
        if (chapter != undefined) oldQuery["chapter"] = chapter;
        if (subject != undefined) oldQuery["subject"] = subject;

        var totalQuestion = allQuestions(oldQuery);
        //alert('Total Question: ' + totalQuestion.count());
        var counter = 0;
        totalQuestion.each(function (r) {
          if ($.inArray(String(r.id), testBank) == -1) {
            counter++;
            {
              var $div = $('<div class="question-parent"></div>').attr({
                id: "fib",
                "data-type": r.type,
                "data-cid": r.id,
              });
              var $checkbox = $("<input/>").attr({
                type: "checkbox",
                id: "check" + counter,
              });

              $checkbox.appendTo($div).click(function (e) {
                var curIndex = testBank.indexOf(String(replacerId));
                //testBank.splice(curIndex, 1);
                //testBank.push(String(r.id));
                testBank[curIndex] = String(r.id);
                $("#replaceQuestion").hide();
                $("#explorer").empty();
                addToTest();
              });
              var $opts;
              for (var q in r.question) {
                console.log(type, " answer");
                if (q == "title") {
                  $('<div class="question-title"></div>')
                    .html(counter + ". " + r.question.title)
                    .appendTo($div);
                } else if (q == "image") {
                  if (r.question.image != undefined && r.question.image != "") {
                    var qimagepathblob = imagePath + r.question.image;
                    if (r.question.image.indexOf("blob:") == 0) {
                      qimagepathblob = r.question.image;
                    }
                    $('<div class="image"></div>')
                      .html('<img src="' + qimagepathblob + '"/>')
                      .appendTo($div);
                  }
                } else if (q == "option") {
                  $opts = $('<div class="options"></div>');
                  if (r.type.toLowerCase() == "mtf") {
                    var optnameii =
                      "<br><table border='1' cellspacing='0' cellpadding='0'>";
                    for (
                      var mtfIndex = 0;
                      mtfIndex < r.question.option.length;
                      mtfIndex++
                    ) {
                      var optnameiis = r.question.option[mtfIndex];
                      var mtfoption = optnameiis.split(" ^ ");

                      optnameii +=
                        "<tr><td>" +
                        mtfoption[0] +
                        "</td><td>" +
                        mtfoption[1] +
                        "</td></tr>";
                    }
                    optnameii += "</table>";

                    $('<div class="option rx-editable"></div>')
                      .html(optnameii)
                      .appendTo($opts);
                  } else {
                    if (r.question.option.length == 1) {
                      var z = 0;
                      var optnameii = r.question.option[z];
                      if (
                        optnameii.indexOf(".jpg") != -1 ||
                        optnameii.indexOf(".png") != -1
                      ) {
                        optnameii =
                          "<img class='see' src='" +
                          imagePath +
                          r.question.option[z] +
                          "'>";
                      }
                      $('<div class="question-title"></div>')
                        .html(optnameii)
                        .appendTo($opts);
                    } else {
                      for (var z = 0; z < r.question.option.length; z++) {
                        //$('<div class="option rx-editable"></div>').html(String.fromCharCode(65+z) + '. ' + r.question.option[z]).appendTo($opts);

                        var optnameii = r.question.option[z];
                        if (
                          optnameii.indexOf(".jpg") != -1 ||
                          optnameii.indexOf(".png") != -1
                        ) {
                          optnameii =
                            "<img class='see' src='" +
                            imagePath +
                            r.question.option[z] +
                            "'>";
                        }
                        //console.log(optnameii);
                        $('<div class="question-title"></div>')
                          .html(optnameii)
                          .appendTo($opts);
                      }
                    }
                  }
                }
              }
              $opts.appendTo($div);
              $('<div class="clear"></div>').appendTo($div);
              $("#replaceList").append($div);
            }
          }
        });
        $("#replaceQuestion").show();
        if ($("#replaceList").is(":empty")) {
          $("#replaceList").html(
            '<div style="text-align: center; line-height: 300px;">No, similar questions available.</div>'
          );
        }
      })
      .appendTo($tools);

    $('<a href="javascript:void(0)"></a>')
      .html("Remove")
      .attr({ class: "tool-remove" })
      .click(function (i, e) {
        var _this = this;

        jQuery.confirm({
          title: "Do you want to delete this question?",
          content: "",
          buttons: {
            confirm: function () {
              var currentId = jQuery(_this).parent().parent().attr("id");

              var curIndex = testBank.indexOf(String(currentId));
              testBank.splice(curIndex, 1);

              console.log(currentId + ": " + curIndex);

              jQuery("div.question-parent")
                .filter(function (i, e) {
                  return jQuery(this).attr("data-cid") == currentId;
                })
                .removeAttr("disabled")
                .prop("checked", false)
                .removeClass("selection");

              addToTest();
            },
            cancel: function () {
              //$.alert('Canceled!');
            },
          },
        });
      })
      .appendTo($tools);

    $('<a href="javascript:void(0)"></a>')
      .html("Edit")
      .attr({ class: "tool-edit" })
      .click(function (e) {
        var _this = this;
        if ($(_this).parents(".question-parent").hasClass("editingmode-on")) {
          var $parent = $(_this).parents(".question-parent");

          $(_this)
            .parents(".question-parent")
            .find(".rx-editable")
            .attr("contenteditable", false);
          $(_this).parents(".question-parent").removeClass("editingmode-on");
          _this.textContent = "Edit";

          var updatedData = {};

          var qTitle = $parent.find(".question-title").html();
          var qOptions = $parent
            .find(".option.rx-editable")
            .map(function (e) {
              return this.innerHTML;
            })
            .get();
          var answer = $parent.find(".simp-ans-text").html();
          var image = $parent.find(".image").html();

          // filter data
          var tmp = document.createElement("div");

          if (image == undefined) {
            image = "";
          } else {
            tmp.innerHTML = image;
            var src = tmp.getElementsByTagName("img")[0].src;
            image = src;
          }

          for (var i = 0; i < qOptions.length; i++) {
            if (qOptions[i].indexOf("<img src") != -1) {
              tmp.innerHTML = qOptions[i];
              var src = tmp.getElementsByTagName("img")[0].src;
              qOptions[i] = src;
            } else {
              qOptions[i] = qOptions[i]
                .split('<span class="aligntop">●</span> ')
                .join("");
            }
          }
          if (answer.indexOf("<img src") != -1) {
            tmp.innerHTML = answer;
            var src = tmp.getElementsByTagName("img")[0].src;
            answer = src;
          }

          var id = parseInt($parent.attr("id"));

          //allQuestions({id: id}).first().question.title = qTitle;
          //allQuestions({id: id}).first().question.option = qOptions;
          allQuestions({ id: id }).first().question.option = [qTitle];
          allQuestions({ id: id }).first().answer = [answer];
        } else {
          $(_this)
            .parents(".question-parent")
            .find(".rx-editable")
            .attr("contenteditable", true);
          $(_this).parents(".question-parent").addClass("editingmode-on");
          _this.textContent = "Save";
          $(_this)
            .parents(".question-parent")
            .find(".rx-editable")
            .first()
            .focus();
        }
      })
      .appendTo($tools);

    $('<div class="clear"></div>').appendTo($div);
    $div.appendTo($container.find("#" + qID));
  });
}

function numValidate(event) {
  var key = window.event ? event.keyCode : event.which;

  if (
    event.keyCode == 8 ||
    event.keyCode == 46 ||
    event.keyCode == 37 ||
    event.keyCode == 39
  ) {
    return true;
  } else if ((key >= 48 && key <= 57) || (key >= 96 && key <= 105)) {
    return true;
  } else return false;
}

function loadImage() {
  var input, file, fr, img;

  if (typeof window.FileReader !== "function") {
    write("The file API isn't supported on this browser yet.");
    return;
  }

  input = document.getElementById("imgfile");
  if (!input) {
    write("Um, couldn't find the imgfile element.");
  } else if (!input.files) {
    write(
      "This browser doesn't seem to support the `files` property of file inputs."
    );
  } else if (!input.files[0]) {
    write("Please select a file before clicking 'Load'");
  } else {
    file = input.files[0];
    fr = new FileReader();
    fr.onload = createImage;
    fr.readAsDataURL(file);
  }

  function createImage() {
    img = new Image();
    img.onload = imageLoaded;
    img.src = fr.result;
  }

  function imageLoaded() {
    var canvas = document.getElementById("canvas");
    canvas.width = img.width;
    canvas.height = img.height;
    var ctx = canvas.getContext("2d");
    ctx.drawImage(img, 0, 0);
    //document.body.innerHTML += '<img src="' + canvas.toDataURL("image/png") + '">';
    userImageLogo =
      '<img class="userDisplayLogos" src="' +
      canvas.toDataURL("image/png") +
      '">';
    $("#previewHTML").prepend(userImageLogo);
  }

  function write(msg) {
    console.log(msg);
  }
}

function setActionForSavingMarks() {
  $("body").on("blur", ".topTitlez .marksInput", function () {
    if (this.value != "" && this.value != 0) {
      var marks = this.value;
      var id = $(this).parents(".question-parent").attr("id");
      allQuestions({ id: parseInt(id) }).first().mark = marks;
    }
  });
}

function addAndMoveOn() {
  //addToTest();
  showCart();
}

function closePopupFromPreview() {
  $("#userInfo").hide();
  resetUserFormPreview();
}

function resetUserFormPreview() {
  $("#level2 #userInfo input[type='text'").val("");
  $("#auto-level-2 #userInfo input[type='text'").val("");

  $("#level2 #userInfo input[type='file'").val("");
  $("#auto-level-2  #userInfo input[type='file'").val("");
}
