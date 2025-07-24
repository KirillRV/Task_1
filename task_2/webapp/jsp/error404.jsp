<%@ page contentType="text/html;charset=UTF-8" language="java" %>
<!DOCTYPE html>
<html>
<head>
    <title>Page Not Found - 404</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/css/bootstrap.min.css" rel="stylesheet">
    <link href="${pageContext.request.contextPath}/css/style.css" rel="stylesheet">
</head>
<body class="container mt-5 text-center">
<h1 class="display-1 text-danger">404</h1>
<h2>Oops! Page Not Found</h2>
<p>The page you are looking for does not exist.</p>
<a href="<%= request.getContextPath() %>/login" class="btn btn-primary">Go to Login Page</a>
</body>
</html>