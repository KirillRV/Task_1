<%@ page contentType="text/html;charset=UTF-8" %>
<%@ taglib prefix="c" uri="http://java.sun.com/jsp/jstl/core" %>
<!DOCTYPE html>
<html>
<head>
    <title>Register</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/css/bootstrap.min.css" rel="stylesheet">
    <link href="${pageContext.request.contextPath}/css/style.css" rel="stylesheet">
</head>
<body class="container mt-5">

<h2>User Registration</h2>

<!-- Display error if exists -->
<c:if test="${not empty error}">
    <div class="alert alert-danger">${error}</div>
</c:if>

<!-- Registration form -->
<form method="post" action="${pageContext.request.contextPath}/register">
    <div class="mb-3">
        <label for="username" class="form-label">Username</label>
        <input type="text" class="form-control" id="username" name="username" required />
    </div>
    <div class="mb-3">
        <label for="password" class="form-label">Password</label>
        <input type="password" class="form-control" id="password" name="password" required />
    </div>
    <button type="submit" class="btn btn-primary">Register</button>
</form>

</body>
</html>