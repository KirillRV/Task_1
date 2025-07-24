<%@ page contentType="text/html;charset=UTF-8" %>
<%@ taglib prefix="c" uri="http://java.sun.com/jsp/jstl/core" %>

<!DOCTYPE html>
<html>
<head>
    <title>Calculator</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/css/bootstrap.min.css" rel="stylesheet">
    <link href="${pageContext.request.contextPath}/css/style.css" rel="stylesheet">
</head>
<body class="container mt-5">

<h2>Welcome, ${sessionScope.user}!</h2>

<a href="${pageContext.request.contextPath}/users" class="btn btn-info mb-3">View All Users</a>

<!-- Calculator form -->
<form method="get" action="${pageContext.request.contextPath}/calc" class="mb-3">
    <div class="row">
        <div class="col">
            <input type="number" name="num1" class="form-control" placeholder="Number 1" step="any" required>
        </div>
        <div class="col">
            <input type="number" name="num2" class="form-control" placeholder="Number 2" step="any" required>
        </div>
        <div class="col">
            <select name="type" class="form-select" required>
                <option value="sum">+</option>
                <option value="sub">−</option>
                <option value="mul">×</option>
                <option value="div">÷</option>
            </select>
        </div>
        <div class="col">
            <button type="submit" class="btn btn-success">Calculate</button>
        </div>
    </div>
</form>

<!-- Show result if exists -->
<c:if test="${not empty result}">
    <div class="alert alert-info">
        Result: ${result}
    </div>
</c:if>

<!-- History of operations -->
<h4>Operation History</h4>
<ul class="list-group">
    <c:forEach var="entry" items="${history}">
        <li class="list-group-item">${entry}</li>
    </c:forEach>
</ul>

<!-- Clear history button -->
<form method="post" action="${pageContext.request.contextPath}/clearHistory" class="mt-3">
    <button type="submit" class="btn btn-danger">Clear History</button>
</form>

</body>
</html>