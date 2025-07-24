<%@ page contentType="text/html;charset=UTF-8" %>
<%@ taglib prefix="c" uri="http://java.sun.com/jsp/jstl/core" %>
<!DOCTYPE html>
<html>
<head>
  <title>Registered Users</title>
  <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/css/bootstrap.min.css" rel="stylesheet">
  <link href="${pageContext.request.contextPath}/css/style.css" rel="stylesheet">
</head>
<body class="container mt-5">

<h2>Registered Users</h2>

<table class="table table-bordered">
  <thead>
  <tr><th>#</th><th>Username</th></tr>
  </thead>
  <tbody>
  <c:forEach var="user" items="${users}" varStatus="loop">
    <tr>
      <td>${loop.index + 1}</td>
      <td>${user}</td>
    </tr>
  </c:forEach>
  </tbody>
</table>

<a href="${pageContext.request.contextPath}/webapp/jsp/calculator.jsp" class="btn btn-secondary">Back to Calculator</a>


</body>
</html>