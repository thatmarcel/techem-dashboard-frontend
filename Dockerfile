FROM oven/bun:latest

WORKDIR /app

COPY --from=builder /app .

EXPOSE 3000

ENTRYPOINT [ "bun", "x", "serve", "-p", "3000" ]
