var $out;
var allQuestions;
$(document).ready(function()
{
	for(var i=0; i<questions.length; i++)
	{
		questions[i]["id"] = i+1;
		$out = $('#treeTemplate');
		for(var j=0; j<treeStructure.length; j++)
		{
			var itemName = questions[i][treeStructure[j]];
			var level = j;
			
			/**********************************************************/
			var $con;
			if($out.children().eq(0).length == 0)
			{
				$con = $('<ul></ul>');
				$con.append('<li>'+ itemName + '</li>');
				$out.append($con);
			}
			else
			{
				var flag = $out.children().eq(0).children().filter(function(i, e)
				{
					 return (this.childNodes[0].textContent == itemName);
				});

				if(flag.length == 0)
					$out.children().eq(0).append('<li>' + itemName + '</li>');
			}
			
			var ele = $out.children(0).eq(0).children().filter(function(i, e)
			{
				return (this.childNodes[0].textContent == itemName);
			}).get();
			
			$out = $(ele);
			/**********************************************************/
		}
	}
	allQuestions = TAFFY(questions);
});