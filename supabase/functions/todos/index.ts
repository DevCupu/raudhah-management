import { withSupabase } from "@supabase/server"

export default {
  fetch: withSupabase({ auth: "user" }, async (_req, ctx) => {
    // ctx.supabase is automatically RLS-scoped to the user's JWT
    // ctx.supabaseAdmin is the service_role client that bypasses RLS
    const { data, error } = await ctx.supabase.from("todos").select()
    
    if (error) {
      return new Response(JSON.stringify({ error: error.message }), {
        status: 500,
        headers: { "Content-Type": "application/json" }
      })
    }

    return Response.json(data)
  }),
}
