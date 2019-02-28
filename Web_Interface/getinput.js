const budgetHTML =
    "<span>" + 
        "<div class='delete'>" + 
            "<div class='backSlash'></div>" + 
            "<div class='forSlash'></div>" + 
        "</div>" + 
        "<input class='budgetCat' type='text' placeholder='Budget Category'>" + 
        "<input class='budget' type='number' placeholder='Budget'>" + 
    "</span>";

$('#budgetCats').on('click', 'div.delete', function() {
    $(this).parent('span').remove();
})

$('#addBudgetCat').on('click', () => {
    $('#budgetCats > div').append(budgetHTML);
})