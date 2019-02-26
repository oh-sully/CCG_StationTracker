$('#budgetCats').on('click', 'div.delete', function() {
    $(this).parent('span').remove()
    console.log("Deleted")
})

$('#addBudgetCat').on('click', () => {
    $('#budgetCats > div').append(
        "<span>" + 
            "<div class='delete'>" + 
                "<div class='backSlash'></div>" + 
                "<div class='forSlash'></div>" + 
            "</div>" + 
            "<input class='budgetCat' type='text'>" + 
            "<input class='budget' type='number'>" + 
        "</span>")
})