import fastify from "fastify";
import { authRoutes } from "./routes/auth";
import { folderRoutes } from "./routes/folders";
import { emailRoutes } from "./routes/emails";

const app = fastify();

app.register(authRoutes, { prefix: "/api/auth" });
app.register(folderRoutes);
app.register(emailRoutes);

app.listen({ port: 3000 });
