// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

function get_template_A_str(user_info)
{
    let str = "";
    const response = await fetch('../Templates/Legence.htm');
    const htmlContent = await response.text();
    str += htmlContent
  if (is_valid_data(user_info.greeting))
  {
    str += user_info.greeting + "<br/>";
  }

  str += "<table>";
  str +=   "<tr>";
  str +=     "<td style='border-right: 1px solid #000000; padding-right: 5px;'><img src='data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAALYAAAAYCAMAAABgOPeoAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAKOUExURWZmZmdnZ2dnZ2dnZwAAAGhoaGpqamlpaWdnZ2NjY2dnZ2lpaWdnZ2hoaGZmZmpqamtra2ZmZmhoaGdnZ2dnZ2hoaGdnZ2ZmZmhoaGpqamZmZmhoaGdnZ2dnZ2lpaW5ubm1tbWhoaGtra2ZmZmpqamNjY2ZmZm9vb2dnZ2NjY2dnZ2pqamhoaGpqamhoaGZmZmlpaWVlZWxsbGhoaGdnZ2lpaWdnZ2VlZWZmZmdnZ2VlZWhoaGhoaGlpaWhoaGhoaGVlZXBwcGZmZmlpaWpqamlpaW5ubmdnZ2hoaG1tbWlpaWZmZmdnZ2lpaWhoaGdnZ2ZmZmdnZ2ZmZmhoaGdnZ2tra21tbWZmZmlpaWdnZ2pqamJiYmRkZGhoaGVlZWNjY21tbW1tbWdnZ2ZmZmZmZmBgYGJiYmVlZWhoaGtra2hoaGdnZ2FhYWhoaGZmZgAAAGZmZmxsbGZmZnJycmZmZmhoaFVVVWdnZ2lpaWhoaG5ubmhoaG1tbWlpaWpqamxsbGhoaGlpaWpqamZmZm5ubm9vb2dnZ2pqamdnZ2dnZ2NjY3R0dHJycnV1dWdnZ2lpaWpqam1tbWdnZ2hoaGpqamVlZW5ubmhoaGtra21tbWpqamlpaWtra2dnZ3BwcGZmZmpqamhoaGRkZGVlZWdnZ2lpaW1tbWdnZ2lpaWhoaG1tbWpqamdnZ2dnZ2dnZ21tbW5ubmhoaGZmZmpqampqamdnZ2hoaHFxcWhoaGZmZmdnZ2VlZWZmZmZmZmhoaGJiYlxcXGdnZ2dnZ2xsbGtra2hoaG1tbWtra2pqamlpaWpqamhoaGlpaWRkZGVlZWxsbGVlZWtra2lpaYCAgGlpaWlpaWlpaV5eXmlpaWpqagjZfMMAAADadFJOU9vt7DQAYvLysxLX7+ssifP0jzHv2hauZ1bx7///OWv//8IT6v8xlv+cNu4YvW9d/zhq/8Dn//0wlJo1u25cbMSX/51wXmb29PP2uP719PIvjpX99vJY9Q8nJSQnHOw6JAcV0l8jEA0m6f/JPhXqGQEtLQodHssDRTNp/Pv8vsX5+Pj4k/39mVvDzzv///+jYU28ubiMbbxxQ7swPSYqEDwp0xcrYREO+F9n+br39pD6+pb5WWPEZf/HFPCaoDfSGhlyxslMU8rKlbTIxscudMt5yZ4Ctsi3E5JICQS0pwAAAAlwSFlzAAAOxAAADsQBlSsOGwAAA4JJREFUWEfdl2l7E1UYhpPpA8giKAhU5pSkOF2k4IFgEKSAodgWqVQp0DSVgmCBsCMU2QvIolhFLOIaUKtsorK3FVAJstgCZVX8N17TmUx63pxMIPgBuL/lec9zrvvTOycOh1NRnM4UxNKuvSKhw2NAx06daawoirPL40DXbjTWeeJJoLvipLGiKD2e6gn06kBjnW69gdSnaWrg6KMypqpp1BlAXxeT4E4H+j2j0ZgxprIMIFNaysoGnlVVGjPGtP45wAA3jXVcA4HnpBPGHArjnLFB1BnAYDeX4BkC9HveS2POORv6AjBMWnINB15kjMacc2+XHGCEh8Y67lxgpHTC+aOnPSpLleAabWmTiVfNAF6SlnxjLO02qf6zjXabSStZecBYF00NbLRfzi8ojGXcKxHt8UXi5NUJxcBr+WJokP96RHti9M6i8aJ2ySShU5A/GJg8ReZQWGijXeovC8RS/oap7Z1aMU2YlAWmA2/6hczE39PU1maUR0rTKqbqt0S0mXfmW0KnzF8JzJotcwgEbLTjY2jPqaS5La3awbnRoHKOqD2v7ekE2GjPT+kuIaUsor1goXV0UaB48eIlS94GlspLyyLaVcut0sIFovY71iRKxxX0KgMb7ZWrVktYs1aivXxd9fr1GzamAe9uogWdzVuS1N66hl5lYKMdZ5O8J9N+362qXm0bkCkt+bKT1E5ikyTY24J2VZBzplYk3Nv3qp3E3n5ItfM+8EioqbLVHiMtffhRktrba+hVBjbaH+/4pDaWHTtttT/dRQs6uz5LUvvzL2QOtbU22vGx0Y5PUtrxMLW/pLkt5uemNJrM/R+0vxKO22Nqp37tFwiVArtDYmYQWhbRdu6xwm++1Uzt7+qE0yZ1OXeh/f1ssRT6AdgrdfD7W7U531dUIlC9HzhwUMwMqn+0nlKHfrLSofrDVdf++RfhtMnhIwm1OT9KHA72BY5Vi1kEU5uuc9fxuJ+bE7KHq/6zVTvBw9VOm5FSwodrLJ7tCfd2DPe3t2NJ+DchlgdeuyioSfCdBHLraarjawAaf3XTWNO0oGcPcEpaqjkNnAkGNS3rt6j274VuTXMX/AGk+2hBp/4ssF860TRHbUNYwrk/gfMXaKpz8RJQ+lcTjcPhcPPlK8DVFhrrtFwDQs0N4fD1Gzct7cZbTeHw7b//Ae5cpAWdC6nAv+doavAfOQIV34F2xRwAAAAASUVORK5CYII=' alt='Logo' /></td>";
  str +=     "<td style='padding-left: 5px;'>";
  str +=	   "<strong>" + user_info.name + "</strong>";
  str +=     is_valid_data(user_info.pronoun) ? "&nbsp;" + user_info.pronoun : "";
  str +=     "<br/>";
  str +=	   is_valid_data(user_info.job) ? user_info.job + "<br/>" : "";
  str +=	   user_info.email + "<br/>";
  str +=	   is_valid_data(user_info.phone) ? user_info.phone + "<br/>" : "";
  str +=     "</td>";
  str +=   "</tr>";
  str += "</table>";

  return str;
}

function get_template_B_str(user_info)
{
  let str = "";
  if (is_valid_data(user_info.greeting))
  {
    str += user_info.greeting + "<br/>";
  }

  str += "<table>";
  str +=   "<tr>";
    str +=     "<td style='border-right: 1px solid #000000; padding-right: 5px;'><img src='data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAALYAAAAYCAMAAABgOPeoAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAKOUExURWZmZmdnZ2dnZ2dnZwAAAGhoaGpqamlpaWdnZ2NjY2dnZ2lpaWdnZ2hoaGZmZmpqamtra2ZmZmhoaGdnZ2dnZ2hoaGdnZ2ZmZmhoaGpqamZmZmhoaGdnZ2dnZ2lpaW5ubm1tbWhoaGtra2ZmZmpqamNjY2ZmZm9vb2dnZ2NjY2dnZ2pqamhoaGpqamhoaGZmZmlpaWVlZWxsbGhoaGdnZ2lpaWdnZ2VlZWZmZmdnZ2VlZWhoaGhoaGlpaWhoaGhoaGVlZXBwcGZmZmlpaWpqamlpaW5ubmdnZ2hoaG1tbWlpaWZmZmdnZ2lpaWhoaGdnZ2ZmZmdnZ2ZmZmhoaGdnZ2tra21tbWZmZmlpaWdnZ2pqamJiYmRkZGhoaGVlZWNjY21tbW1tbWdnZ2ZmZmZmZmBgYGJiYmVlZWhoaGtra2hoaGdnZ2FhYWhoaGZmZgAAAGZmZmxsbGZmZnJycmZmZmhoaFVVVWdnZ2lpaWhoaG5ubmhoaG1tbWlpaWpqamxsbGhoaGlpaWpqamZmZm5ubm9vb2dnZ2pqamdnZ2dnZ2NjY3R0dHJycnV1dWdnZ2lpaWpqam1tbWdnZ2hoaGpqamVlZW5ubmhoaGtra21tbWpqamlpaWtra2dnZ3BwcGZmZmpqamhoaGRkZGVlZWdnZ2lpaW1tbWdnZ2lpaWhoaG1tbWpqamdnZ2dnZ2dnZ21tbW5ubmhoaGZmZmpqampqamdnZ2hoaHFxcWhoaGZmZmdnZ2VlZWZmZmZmZmhoaGJiYlxcXGdnZ2dnZ2xsbGtra2hoaG1tbWtra2pqamlpaWpqamhoaGlpaWRkZGVlZWxsbGVlZWtra2lpaYCAgGlpaWlpaWlpaV5eXmlpaWpqagjZfMMAAADadFJOU9vt7DQAYvLysxLX7+ssifP0jzHv2hauZ1bx7///OWv//8IT6v8xlv+cNu4YvW9d/zhq/8Dn//0wlJo1u25cbMSX/51wXmb29PP2uP719PIvjpX99vJY9Q8nJSQnHOw6JAcV0l8jEA0m6f/JPhXqGQEtLQodHssDRTNp/Pv8vsX5+Pj4k/39mVvDzzv///+jYU28ubiMbbxxQ7swPSYqEDwp0xcrYREO+F9n+br39pD6+pb5WWPEZf/HFPCaoDfSGhlyxslMU8rKlbTIxscudMt5yZ4Ctsi3E5JICQS0pwAAAAlwSFlzAAAOxAAADsQBlSsOGwAAA4JJREFUWEfdl2l7E1UYhpPpA8giKAhU5pSkOF2k4IFgEKSAodgWqVQp0DSVgmCBsCMU2QvIolhFLOIaUKtsorK3FVAJstgCZVX8N17TmUx63pxMIPgBuL/lec9zrvvTOycOh1NRnM4UxNKuvSKhw2NAx06daawoirPL40DXbjTWeeJJoLvipLGiKD2e6gn06kBjnW69gdSnaWrg6KMypqpp1BlAXxeT4E4H+j2j0ZgxprIMIFNaysoGnlVVGjPGtP45wAA3jXVcA4HnpBPGHArjnLFB1BnAYDeX4BkC9HveS2POORv6AjBMWnINB15kjMacc2+XHGCEh8Y67lxgpHTC+aOnPSpLleAabWmTiVfNAF6SlnxjLO02qf6zjXabSStZecBYF00NbLRfzi8ojGXcKxHt8UXi5NUJxcBr+WJokP96RHti9M6i8aJ2ySShU5A/GJg8ReZQWGijXeovC8RS/oap7Z1aMU2YlAWmA2/6hczE39PU1maUR0rTKqbqt0S0mXfmW0KnzF8JzJotcwgEbLTjY2jPqaS5La3awbnRoHKOqD2v7ekE2GjPT+kuIaUsor1goXV0UaB48eIlS94GlspLyyLaVcut0sIFovY71iRKxxX0KgMb7ZWrVktYs1aivXxd9fr1GzamAe9uogWdzVuS1N66hl5lYKMdZ5O8J9N+362qXm0bkCkt+bKT1E5ikyTY24J2VZBzplYk3Nv3qp3E3n5ItfM+8EioqbLVHiMtffhRktrba+hVBjbaH+/4pDaWHTtttT/dRQs6uz5LUvvzL2QOtbU22vGx0Y5PUtrxMLW/pLkt5uemNJrM/R+0vxKO22Nqp37tFwiVArtDYmYQWhbRdu6xwm++1Uzt7+qE0yZ1OXeh/f1ssRT6AdgrdfD7W7U531dUIlC9HzhwUMwMqn+0nlKHfrLSofrDVdf++RfhtMnhIwm1OT9KHA72BY5Vi1kEU5uuc9fxuJ+bE7KHq/6zVTvBw9VOm5FSwodrLJ7tCfd2DPe3t2NJ+DchlgdeuyioSfCdBHLraarjawAaf3XTWNO0oGcPcEpaqjkNnAkGNS3rt6j274VuTXMX/AGk+2hBp/4ssF860TRHbUNYwrk/gfMXaKpz8RJQ+lcTjcPhcPPlK8DVFhrrtFwDQs0N4fD1Gzct7cZbTeHw7b//Ae5cpAWdC6nAv+doavAfOQIV34F2xRwAAAAASUVORK5CYII=' alt='Logo' /></td>";
  str +=     "<td style='padding-left: 5px;'>";
  str +=	   "<strong>" + user_info.name + "</strong>";
  str +=     is_valid_data(user_info.pronoun) ? "&nbsp;" + user_info.pronoun : "";
  str +=     "<br/>";
  str +=	   user_info.email + "<br/>";
  str +=	   is_valid_data(user_info.phone) ? user_info.phone + "<br/>" : "";
  str +=     "</td>";
  str +=   "</tr>";
  str += "</table>";

  return str;
}

function get_template_C_str(user_info)
{
  let str = "";
  if (is_valid_data(user_info.greeting))
  {
    str += user_info.greeting + "<br/>";
  }

    str += user_info.name;
    str += "Test"
  
  return str;
}